function new-placemark {
param($name, $lat, $long)

  $point = new-object SharpKml.Dom.Point
  $point.Coordinate = new-object SharpKml.Base.Vector $lat, $long
  $place = new-object SharpKml.Dom.Placemark
  $place.Geometry = $point
  $place.Name = $name
  return $place
}


$scriptPath = split-path -parent $MyInvocation.MyCommand.Definition

add-type -path "$scriptPath\epplus.dll"
add-type -path "$scriptPath\SharpKml.dll"

$xl = new-object OfficeOpenXml.ExcelPackage "$scriptPath\wa-place-names.xlsx"

$sheet = $xl.Workbook.WorkSheets[1]
$lastRow = $sheet.Dimension.End.Row


$columns = @{}
1..$sheet.Dimension.End.Column | %{$columns[$sheet.Cells[1,$_].Value] = $_}

$folders = @{}
2..$lastRow | % {
  $lineNumber = $_
  $countyName = $sheet.Cells[$_, $columns['COUNTY_NAME']].Value
  if ([string]::IsNullOrWhitespace($countyName)) { $countyName = "Other" }
  
  try {
  $folder = $folders[$countyName]
  if ($folder -eq $null) {
    $folder = new-object SharpKml.Dom.Folder
    $folder.Name = $countyName
    $folders[$countyName] = $folder
  }
  
  $feature = new-placemark $sheet.Cells[$_, $columns['FEATURE_NAME']].Value $sheet.Cells[$_, $columns['PRIM_LAT_DEC']].Value $sheet.Cells[$_, $columns['PRIM_LONG_DEC']].Value
  $folder.AddFeature($feature)
  }
   catch
   {
   write-error "Error on line $lineNumber"
   }
}

$parentFolder = new-object SharpKml.Dom.Folder
$folders.Keys | sort-object | %{
  $parentFolder.AddFeature($folders[$_])
}

$kml = new-object SharpKml.Dom.Kml
$kml.Feature = $parentFolder

$s = new-object SharpKml.Base.Serializer
$s.Serialize($kml)
$s.Xml > test.kml