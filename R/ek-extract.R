library(ISOcodes)
library(openxlsx)
library(SpectrumUtils)

process_pjnzs = function(pjnz_list) {
  data_list = lapply(pjnz_list, extract_pjnz_data)
  
  wb = createWorkbook()
  addtab(wb, prepare_frame_all(data_list, "diag_all"), "CaseTot")
  addtab(wb, prepare_frame_sex(data_list, "diag_sex"), "CaseM-F")
  addtab(wb, prepare_frame_age(data_list, "diag_age"), "CaseAge")
  addtab(wb, prepare_frame_all(data_list, "mort_all"), "DeathTot")
  addtab(wb, prepare_frame_sex(data_list, "mort_sex"), "DeathM-F")
  addtab(wb, prepare_frame_age(data_list, "mort_age"), "DeathAge")
  addtab(wb, prepare_frame_cd4(data_list, "diag_cd4"), "CD4")
  addtab(wb, prepare_frame_age(data_list, "migr"), "ImmigrPrevPos")
  addtab(wb, prepare_frame_meta(data_list, check_incidence), "IncidTrendRule")
  addtab(wb, prepare_frame_meta(data_list, check_irr_state), "SexAgeIRR")
  addtab(wb, prepare_frame_meta(data_list, check_knowledge), "KOSTrendRule")
  return(wb)
}

country_name = function(iso3) {
  isostr = sprintf("%03d", as.numeric(iso3))
  cnames = ISO_3166_1$Name[match(isostr, ISO_3166_1$Numeric)]
  return(cnames)
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
    pjnz = proj_name, # pjnz_name,
    country = country_name(geo_info$iso.code),
    isocode = geo_info$iso.code,
    snuname = geo_info$snu.name,

    inci_model = dp.inputs.incidence.model(dp),
    inci_curve = dp.inputs.csavr.model(dp),
    csavr_irrs = dp.inputs.csavr.irr.options(dp, direction="wide"),

    irr_custom  = dp.inputs.irr.custom(dp),
    irr_pattern = dp.inputs.irr.pattern(dp),

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
  addWorksheet(workbook, tabname)
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
  return(dplyr::left_join(meta, data_wide, by=c("PJNZ")))
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
             "AIM IRR custom selected?")
  irr_opts = dat$csavr_irrs[dat$csavr_irrs$Model == dat$inci_curve,]

  rval = data.frame(PJNZ    = dat$pjnz,
                    Model   = dat$inci_model,
                    Curve   = dat$inci_curve,
                    SexIRRs = irr_opts$Sex[1],
                    AgeIRRs = irr_opts$Age[1],
                    AIMPattern = dat$irr_pattern,
                    AIMCustom  = dat$irr_custom)
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


