library(ISOcodes)
library(openxlsx)
library(SpectrumUtils)

process_pjnzs = function(pjnz_list) {
  data_list = lapply(pjnz_list, extract_pjnz_data)
  
  fm_diag_all = prepare_frame_all(data_list, "diag_all")
  fm_diag_sex = prepare_frame_sex(data_list, "diag_sex")
  fm_diag_age = prepare_frame_age(data_list, "diag_age")
  fm_mort_all = prepare_frame_all(data_list, "mort_all")
  fm_mort_sex = prepare_frame_sex(data_list, "mort_sex")
  fm_mort_age = prepare_frame_age(data_list, "mort_age")
  
  wb = createWorkbook()
  addtab(wb, fm_diag_all, "CaseTot")
  addtab(wb, fm_diag_sex, "CaseM-F")
  addtab(wb, fm_diag_age, "CaseAge")
  addtab(wb, fm_mort_all, "DeathTot")
  addtab(wb, fm_mort_sex, "DeathM-F")
  addtab(wb, fm_mort_age, "DeathAge")
  # addtab(wb, prepare_frame_cd4(data_list, "diag_cd4"), "CD4")
  # addtab(wb, prepare_frame_age(data_list, "migr"), "ImmigrPrevPos")
  # 
  # inci.tname = "IncidTrendRule"
  # inci.frame = prepare_frame_meta(data_list, check_incidence)
  # addtab(wb, inci.frame, inci.tname)
  # ci = 1 + ncol(inci.frame)
  # for (ri in 1:nrow(inci.frame)) {
  #   expr = sprintf("AND(%s%d>=8,%s%d>=2019)", int2col(ci-2), ri+1, int2col(ci-1), ri+1)
  #   writeFormula(wb, sheet=inci.tname, x=expr, startCol=ci, startRow=ri+1)
  # }
  # writeData(wb, sheet=inci.tname, x="Meets the rule for a UNAIDS-publishable 2010-2021 incidence trend:", startCol=ci, startRow=1)
  # 
  # addtab(wb, prepare_frame_meta(data_list, check_irr_state), "SexAgeIRR")
  # 
  # know.tname = "KOSTrendRule"
  # know.frame = prepare_frame_meta(data_list, check_knowledge)
  # addtab(wb, know.frame, know.tname)
  # ci = 1 + ncol(know.frame)
  # for (ri in 1:nrow(know.frame)) {
  #   expr = sprintf("AND(%s%d>=1,%s%d)", int2col(ci-2), ri+1, int2col(ci-1), ri+1)
  #   writeFormula(wb, sheet=know.tname, x=expr, startCol=ci, startRow=ri+1)
  # }
  # writeData(wb, sheet=know.tname, x="Meets the rule for a UNAIDS-publishable 2010-2021 Knowledge-of-Status trend:", startCol=ci, startRow=1)
  
  add_crosscheck_sex(wb, "CheckCaseM-F",  "CaseTot",  "CaseM-F",  fm_diag_all, fm_diag_sex)
  add_crosscheck_age(wb, "CheckCaseAge",  "CaseM-F",  "CaseAge",  fm_diag_sex, fm_diag_age)
  add_crosscheck_sex(wb, "CheckDeathM-F", "DeathTot", "DeathM-F", fm_mort_all, fm_mort_sex)
  add_crosscheck_age(wb, "CheckDeathAge", "DeathM-F", "DeathAge", fm_mort_sex, fm_mort_age)
  
  return(wb)
}

country_name = function(iso3) {
  isostr = sprintf("%03d", as.numeric(iso3))
  cnames = ISO_3166_1$Name[match(isostr, ISO_3166_1$Numeric)]
  return(cnames)
}

country_alpha = function(iso3) {
  isostr = sprintf("%03d", as.numeric(iso3))
  codes = ISO_3166_1$Alpha_3[match(isostr, ISO_3166_1$Numeric)]
  return(codes)
}

extract_pjnz_data = function(pjnz_full) {
  pjnz_name = basename(pjnz_full)
  pj = read.raw.pjn(pjnz_full)
  dp = read.raw.dp(pjnz_full)
  yr_bgn = dp.inputs.first.year(dp)
  yr_end = dp.inputs.final.year(dp)

  geo_info = extract.geo.info(pj)
  proj_name = extract.proj.name(pj)

  rval = list(
    pjnz = proj_name,
    country = country_name(geo_info$iso.code),
    isocode = country_alpha(geo_info$iso.code),
    snuname = geo_info$snu.name,

    inci_model = dp.inputs.incidence.model(dp),
    inci_curve = dp.inputs.csavr.model(dp),
    csavr_irrs = dp.inputs.csavr.irr.options(dp, direction="wide"),

    irr_custom  = dp.inputs.irr.custom(dp),
    irr_pattern = dp.inputs.irr.pattern(dp),
    irr_epp     = dp.inputs.irr.sex.from.epp(dp),

    diag_all = dp.inputs.csavr.diagnoses(dp, direction="long", first.year=yr_bgn, final.year=yr_end),
    diag_sex = dp.inputs.csavr.diagnoses.sex(dp, direction="long", first.year=yr_bgn, final.year=yr_end),
    diag_age = dp.inputs.csavr.diagnoses.sex.age(dp, direction="long", first.year=yr_bgn, final.year=yr_end),
    diag_cd4 = dp.inputs.csavr.diagnoses.cd4(dp, direction="long", first.year=yr_bgn, final.year=yr_end),

    mort_all = dp.inputs.csavr.deaths(dp, direction="long", first.year=yr_bgn, final.year=yr_end),
    mort_sex = dp.inputs.csavr.deaths.sex(dp, direction="long", first.year=yr_bgn, final.year=yr_end),
    mort_age = dp.inputs.csavr.deaths.sex.age(dp, direction="long", first.year=yr_bgn, final.year=yr_end),

    migr = dp.inputs.csavr.migr.diagnoses(dp, direction="long", first.year=yr_bgn, final.year=yr_end)
  )
  return(rval)
}

prepare_metadata = function(pjnz_data) {
  meta_list = lapply(pjnz_data, function(pjnz_item) {
    data.frame(
      PJNZ    = pjnz_item$pjnz,
      ISO3    = pjnz_item$isocode,
      Country = pjnz_item$country,
      SNU     = pjnz_item$snuname)
  })
  return(dplyr::bind_rows(meta_list))
}

prepare_flatdata = function(pjnz_data, item_name) {
  list_name = sapply(pjnz_data, function(pjnz_item) {pjnz_item$pjnz})
  list_data = lapply(pjnz_data, function(pjnz_item) {pjnz_item[[item_name]]})
  names(list_data) = list_name
  return(dplyr::bind_rows(list_data, .id="PJNZ"))
}

addtab = function(workbook, tabdata, tabname) {
  ## The header style is applied to twice as many columns as are present in
  ## tabdata. This is a workaround because (1) we may add more columns later and
  ## (2) openxlsx currently has no way to style an entire row or column
  headerStyle = createStyle(wrapText=TRUE, textDecoration="bold")
  addWorksheet(workbook, tabname)
  addStyle(workbook, sheet=tabname, headerStyle, rows=1, cols=1:(2*ncol(tabdata)))
  writeData(workbook, tabname, tabdata)
}

prepare_frame_all = function(pjnz_data, indicator) {
  meta = prepare_metadata(pjnz_data)
  meta$Sex = "Male+Female"
  meta$Age = "15+"
  data_long = prepare_flatdata(pjnz_data, indicator)
  data_wide = reshape2::dcast(data_long, PJNZ~Year, value.var="Value")
  return(dplyr::left_join(meta, data_wide, by=c("PJNZ")))
}

prepare_frame_sex = function(pjnz_data, indicator) {
  meta = prepare_metadata(pjnz_data)
  meta$Age = "15+"
  data_long = prepare_flatdata(pjnz_data, indicator)
  data_wide = reshape2::dcast(data_long, PJNZ+Sex~Year, value.var="Value")
  data_join = dplyr::left_join(meta, data_wide, by=c("PJNZ"))
  
  ## standardize column ordering
  data_join[,1:6] = data_join[,c("PJNZ", "ISO3", "Country", "SNU", "Sex", "Age")]
  
  return(data_join)
}

prepare_frame_age = function(pjnz_data, indicator) {
  meta = prepare_metadata(pjnz_data)
  data_long = prepare_flatdata(pjnz_data, indicator)
  data_wide = reshape2::dcast(data_long, PJNZ+Sex+Age~Year, value.var="Value")
  return(dplyr::left_join(meta, data_wide, by=c("PJNZ")))
}

prepare_frame_cd4 = function(pjnz_data, indicator) {
  meta = prepare_metadata(pjnz_data)
  meta$Sex = "Male+Female"
  data_long = prepare_flatdata(pjnz_data, indicator)
  data_wide = reshape2::dcast(data_long, PJNZ+CD4~Year, value.var="Value")
  return(dplyr::left_join(meta, data_wide, by=c("PJNZ")))
}

## Check if there is sufficient data to estimate incidence
check_incidence = function(dat) {
  first_year = 1990
  final_year = 2021
  cnames = c("PJNZ",
             sprintf("Years of death data during %s-%s", first_year, final_year),
             sprintf("Latest death data data available during %s-%s", first_year, final_year))
  
  sset = subset(dat$mort_all, is.finite(Value) & Year >= first_year & Year <= final_year)
  rval = data.frame(PJNZ  = dat$pjnz,
                    Years = nrow(sset),
                    Final = max(sset$Year))
  colnames(rval) = cnames
  return(rval)
}

## Check if there is sufficient data to estimate knowledge of status
check_knowledge = function(dat) {
  first_year = 2019
  final_year = 2021
  cnames = c("PJNZ",
             sprintf("Years of deaths data during %s-%s", first_year, final_year),
             "New diagnoses entered in 2019?")
  
  sset_mort = subset(dat$mort_all, is.finite(Value) & Year >= first_year & Year <= final_year)
  sset_diag = subset(dat$diag_all, is.finite(Value) & Year == 2019)
  rval = data.frame(PJNZ = dat$pjnz,
                    Years = nrow(sset_mort),
                    Diag  = nrow(sset_diag) > 0)
  colnames(rval) = cnames
  return(rval)
}

check_irr_state = function(dat) {
  cnames = c("PJNZ",
             "Incidence option",
             "CSAVR model",
             "CSAVR Sex IRRs selected?",
             "CSAVR Age IRRs selected?",
             "AIM IRR pattern",
             "AIM IRR custom selected?",
             "AIM Sex IRRs from EPP or AEM?")
  irr_opts = dat$csavr_irrs[dat$csavr_irrs$Model == dat$inci_curve,]

  rval = data.frame(PJNZ    = dat$pjnz,
                    Model   = dat$inci_model,
                    Curve   = dat$inci_curve,
                    SexIRRs = irr_opts$Sex[1],
                    AgeIRRs = irr_opts$Age[1],
                    AIMPattern = dat$irr_pattern,
                    AIMCustom  = dat$irr_custom,
                    SexIRREPP  = dat$irr_epp)
  colnames(rval) = cnames
  return(rval)
}

## fn is a function. fn(pjnz_data[[i]]) must return a data frame with one row;
## one column must be named "PJNZ" and store the PJNZ name
prepare_frame_meta = function(pjnz_data, fn) {
  meta = prepare_metadata(pjnz_data)
  func_list = lapply(pjnz_data, fn)
  func_flat = dplyr::bind_rows(func_list)
  return(dplyr::left_join(meta, func_flat, by=c("PJNZ")))
}

## Add a tab to the workbook that compares overall numbers of diagnoses or
## deaths to data entered by sex. This does not do the comparison, but rather
## produces formulas that do the comparison in Excel.
## 
## tab_name New tab name
## tab_base Tab with baseline data
## tab_comp Tab with comparison data
## dat_base Data frame of baseline data, used for establishing formula ranges
## dat_comp Data frame of comparison data, used for establishing formula ranges
add_crosscheck_sex = function(workbook, tab_name, tab_base, tab_comp, dat_base, dat_comp) {
  meta = dat_base[,1:6]
  cols_meta = ncol(meta)
  
  ## temporary data to establish column names and header style
  temp = dat_base[,(cols_meta+1):ncol(dat_base)]
  temp[1:nrow(temp),1:ncol(temp)] = NA
  
  ## Add a tab with placeholder (missing) data
  addtab(workbook, cbind(meta, temp), tabname=tab_name)
  
  ## Write formulas. Not idiomatic R, but the code is unreadable enough without
  ## vectorizing it
  for (ri in 1:nrow(dat_base)) {
    for (ci in 1:ncol(temp)) {
      check_pjnz = sprintf("$A%d", ri+1)
      
      ## Create a formula that sums values from tab_base columns
      range_base_pjnz = sprintf("'%s'!$A$2:$A$%d", tab_base, nrow(dat_base)+1)
      range_base_vals = sprintf("'%s'!%s$2:%s$%d", tab_base, int2col(ci+cols_meta), int2col(ci+cols_meta), nrow(dat_base)+1)
      value_base = sprintf("SUMIF(%s,%s,%s)", range_base_pjnz, check_pjnz, range_base_vals)
      
      ## Create a formula that sums values from tab_comp columns
      range_comp_pjnz = sprintf("'%s'!$A$2:$A$%d", tab_comp, nrow(dat_comp)+1)
      range_comp_vals = sprintf("'%s'!%s$2:%s$%d", tab_comp, int2col(ci+cols_meta), int2col(ci+cols_meta), nrow(dat_comp)+1)
      value_comp = sprintf("SUMIF(%s,%s,%s)", range_comp_pjnz, check_pjnz, range_comp_vals)
      
      expr = sprintf("IF(%s>0,%s/%s,\"\")", value_comp, value_comp, value_base)
      
      writeFormula(workbook, sheet=tab_name, x=expr, startCol=ci+ncol(meta), startRow=ri+1)
    }
  }
}

## Add a tab to the workbook that compares numbers of diagnoses or deaths by sex
## and age to data entered by sex. This does not do the comparison, but rather
## produces formulas that do the comparison in Excel.
## 
## tab_name New tab name
## tab_base Tab with baseline data
## tab_comp Tab with comparison data
## dat_base Data frame of baseline data, used for establishing formula ranges
## dat_comp Data frame of comparison data, used for establishing formula ranges
add_crosscheck_age = function(workbook, tab_name, tab_base, tab_comp, dat_base, dat_comp) {
  meta = dat_base[,1:6]
  cols_meta = ncol(meta)
  
  ## temporary data to establish column names and header style
  temp = dat_base[,(cols_meta+1):ncol(dat_base)]
  temp[1:nrow(temp),1:ncol(temp)] = NA
  
  ## Add a tab with placeholder (missing) data
  addtab(workbook, cbind(meta, temp), tabname=tab_name)
  
  ## Write formulas. Not idiomatic R, but the code is unreadable enough without
  ## vectorizing it
  for (ri in 1:nrow(dat_base)) {
    for (ci in 1:ncol(temp)) {
      check_pjnz = sprintf("$A%d", ri+1)
      check_sex  = sprintf("$E%d", ri+1)
      
      ## Create a formula that sums values from tab_base columns
      range_base_pjnz = sprintf("'%s'!$A$2:$A$%d", tab_base, nrow(dat_base)+1)
      range_base_sex  = sprintf("'%s'!$E$2:$E$%d", tab_base, nrow(dat_base)+1)
      range_base_vals = sprintf("'%s'!%s$2:%s$%d", tab_base, int2col(ci+cols_meta), int2col(ci+cols_meta), nrow(dat_base)+1)
      value_base = sprintf("SUMIFS(%s,%s,%s,%s,%s)", range_base_vals, range_base_pjnz, check_pjnz, range_base_sex, check_sex)
      
      ## Create a formula that sums values from tab_comp columns
      range_comp_pjnz = sprintf("'%s'!$A$2:$A$%d", tab_comp, nrow(dat_comp)+1)
      range_comp_sex  = sprintf("'%s'!$E$2:$E$%d", tab_comp, nrow(dat_comp)+1)
      range_comp_vals = sprintf("'%s'!%s$2:%s$%d", tab_comp, int2col(ci+cols_meta), int2col(ci+cols_meta), nrow(dat_comp)+1)
      value_comp = sprintf("SUMIFS(%s,%s,%s,%s,%s)", range_comp_vals, range_comp_pjnz, check_pjnz, range_comp_sex, check_sex)
      
      expr = sprintf("IF(%s>0,%s/%s,\"\")", value_comp, value_comp, value_base)
      
      writeFormula(workbook, sheet=tab_name, x=expr, startCol=ci+ncol(meta), startRow=ri+1)
    }
  }
}

