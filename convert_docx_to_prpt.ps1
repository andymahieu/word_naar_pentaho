# convert_docx_to_prpt.ps1
# Generates a valid Pentaho 5.4 Report Template (.prpt) from Stortformulier.docx
# IMPORTANT: Uses ZipArchive directly so that:
#   1. 'mimetype' is always the FIRST entry and stored WITHOUT compression (required by ODF spec)
#   2. All XML files are written UTF-8 WITHOUT BOM (avoids SAXParseException "Content not allowed in prolog")

Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

$ErrorActionPreference = "Stop"
$prptPath = Join-Path (Get-Location) "Stortformulier.prpt"

# Helper: write a string into a ZipArchive entry WITHOUT BOM, using raw bytes
function Add-ZipEntry {
    param(
        [System.IO.Compression.ZipArchive]$Archive,
        [string]$EntryName,
        [string]$Content,
        [System.IO.Compression.CompressionLevel]$Compression = [System.IO.Compression.CompressionLevel]::Optimal
    )
    $entry = $Archive.CreateEntry($EntryName, $Compression)
    $stream = $entry.Open()
    # UTF-8 WITHOUT BOM – write raw bytes directly
    $encoding = New-Object System.Text.UTF8Encoding $false
    $bytes = $encoding.GetBytes($Content)
    $stream.Write($bytes, 0, $bytes.Length)
    $stream.Flush()
    $stream.Close()
}

Write-Host "Generating Stortformulier.prpt..."

if (Test-Path $prptPath) { Remove-Item $prptPath -Force }

$fileStream = [System.IO.File]::Open($prptPath, [System.IO.FileMode]::Create)
$zip = New-Object System.IO.Compression.ZipArchive($fileStream, [System.IO.Compression.ZipArchiveMode]::Create)

# -----------------------------------------------------------------------
# 1. mimetype  –  MUST be first, MUST use NoCompression (ODF spec requirement)
#    Written as raw ASCII bytes directly – no StreamWriter, no BOM, no deflate
# -----------------------------------------------------------------------
$mimetypeEntry = $zip.CreateEntry("mimetype", [System.IO.Compression.CompressionLevel]::NoCompression)
$mimetypeStream = $mimetypeEntry.Open()
$mimetypeBytes = [System.Text.Encoding]::ASCII.GetBytes("application/vnd.pentaho.reporting-archive")
$mimetypeStream.Write($mimetypeBytes, 0, $mimetypeBytes.Length)
$mimetypeStream.Flush()
$mimetypeStream.Close()

# -----------------------------------------------------------------------
# 2. META-INF/manifest.xml
# -----------------------------------------------------------------------
$manifest = '<?xml version="1.0" encoding="UTF-8"?>
<manifest xmlns="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0">
  <file-entry media-type="application/vnd.pentaho.reporting-archive" full-path="/"/>
  <file-entry media-type="text/xml" full-path="layout.xml"/>
  <file-entry media-type="text/xml" full-path="styles.xml"/>
  <file-entry media-type="text/xml" full-path="datadefinition.xml"/>
  <file-entry media-type="text/xml" full-path="settings.xml"/>
  <file-entry media-type="text/xml" full-path="meta.xml"/>
</manifest>'
Add-ZipEntry $zip "META-INF/manifest.xml" $manifest

# -----------------------------------------------------------------------
# 3. meta.xml
# -----------------------------------------------------------------------
$meta = '<?xml version="1.0" encoding="UTF-8"?>
<office:document-meta office:version="1.0"
    xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0"
    xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0">
  <office:meta>
    <meta:generator>Pentaho Report Designer 5.4</meta:generator>
    <meta:title>Stortformulier</meta:title>
    <meta:description>Identificatieformulier Niet-Gevaarlijke Afvalstoffen</meta:description>
  </office:meta>
</office:document-meta>'
Add-ZipEntry $zip "meta.xml" $meta

# -----------------------------------------------------------------------
# 4. settings.xml
# -----------------------------------------------------------------------
$settings = '<?xml version="1.0" encoding="UTF-8"?>
<settings xmlns="http://reporting.pentaho.org/namespaces/engine/classic/bundle/settings/1.0">
  <configuration/>
</settings>'
Add-ZipEntry $zip "settings.xml" $settings

# -----------------------------------------------------------------------
# 5. datadefinition.xml  – all 20 form parameters
# -----------------------------------------------------------------------
$data = '<?xml version="1.0" encoding="UTF-8"?>
<data-definition xmlns="http://reporting.pentaho.org/namespaces/engine/classic/bundle/data/1.0">
  <parameter-definition>
    <parameter name="AfvaltransportAanvoer__afvoer" class="java.lang.String" value-type="java.lang.String"/>
    <parameter name="WerknemerNaam"              class="java.lang.String" value-type="java.lang.String"/>
    <parameter name="WerknemerVoornaam"          class="java.lang.String" value-type="java.lang.String"/>
    <parameter name="VoertuigenNummerplaat"      class="java.lang.String" value-type="java.lang.String"/>
    <parameter name="WerkficheCode_1"            class="java.lang.String" value-type="java.lang.String"/>
    <parameter name="WerfWerfnaam"               class="java.lang.String" value-type="java.lang.String"/>
    <parameter name="WerfWerfadresAdres_1"       class="java.lang.String" value-type="java.lang.String"/>
    <parameter name="WerfWerfadresPostcode"      class="java.lang.String" value-type="java.lang.String"/>
    <parameter name="WerfWerfadresGemeente"      class="java.lang.String" value-type="java.lang.String"/>
    <parameter name="WerfWerfnr"                 class="java.lang.String" value-type="java.lang.String"/>
    <parameter name="WerfProjectProjectnr"       class="java.lang.String" value-type="java.lang.String"/>
    <parameter name="VerwerkersBedrijfsnaam"     class="java.lang.String" value-type="java.lang.String"/>
    <parameter name="VerwerkersStraat"           class="java.lang.String" value-type="java.lang.String"/>
    <parameter name="VerwerkersPostcode"         class="java.lang.String" value-type="java.lang.String"/>
    <parameter name="VerwerkersGemeente"         class="java.lang.String" value-type="java.lang.String"/>
    <parameter name="AfvaltransportRecyclage"    class="java.lang.String" value-type="java.lang.String"/>
    <parameter name="Euralcode"                  class="java.lang.String" value-type="java.lang.String"/>
    <parameter name="WerkficheNr"               class="java.lang.String" value-type="java.lang.String"/>
    <parameter name="WerkficheNaam"             class="java.lang.String" value-type="java.lang.String"/>
    <parameter name="EventStart_Day"            class="java.lang.String" value-type="java.lang.String"/>
  </parameter-definition>
</data-definition>'
Add-ZipEntry $zip "datadefinition.xml" $data

# -----------------------------------------------------------------------
# 6. styles.xml  – A4 portrait, 36pt margins
# -----------------------------------------------------------------------
$styles = '<?xml version="1.0" encoding="UTF-8"?>
<style-sheet xmlns="http://reporting.pentaho.org/namespaces/engine/classic/bundle/style/1.0">
  <page-definition>
    <page-format width="595.275590551181" height="841.8897637795276"
                 orientation="portrait"
                 margin-left="36.0" margin-right="36.0"
                 margin-top="36.0"  margin-bottom="36.0"/>
  </page-definition>
</style-sheet>'
Add-ZipEntry $zip "styles.xml" $styles

# -----------------------------------------------------------------------
# 7. layout.xml  – premium absolute-positioned form layout
# -----------------------------------------------------------------------
$layout = '<?xml version="1.0" encoding="UTF-8"?>
<layout xmlns="http://reporting.pentaho.org/namespaces/engine/classic/bundle/layout/1.0"
        xmlns:core="http://reporting.pentaho.org/namespaces/engine/attributes/core"
        xmlns:style="http://reporting.pentaho.org/namespaces/engine/classic/bundle/style/1.0">

  <report-header>

    <!-- ======================================================= -->
    <!-- TITLE BANNER                                             -->
    <!-- ======================================================= -->
    <rectangle>
      <style:element-style>
        <style:key name="x">0.0</style:key>      <style:key name="y">0.0</style:key>
        <style:key name="width">523.0</style:key> <style:key name="height">40.0</style:key>
        <style:key name="fill-color">#1E3A8A</style:key>
        <style:key name="fill-shape">true</style:key>
        <style:key name="draw-shape">false</style:key>
      </style:element-style>
    </rectangle>
    <label>
      <style:element-style>
        <style:key name="x">10.0</style:key>      <style:key name="y">8.0</style:key>
        <style:key name="width">503.0</style:key> <style:key name="height">24.0</style:key>
        <style:key name="font-name">Arial</style:key>
        <style:key name="font-size">12</style:key>
        <style:key name="font-bold">true</style:key>
        <style:key name="text-color">#FFFFFF</style:key>
        <style:key name="alignment">center</style:key>
      </style:element-style>
      <core:value>IDENTIFICATIEFORMULIER NIET-GEVAARLIJKE AFVALSTOFFEN</core:value>
    </label>

    <!-- Indicator row -->
    <label>
      <style:element-style>
        <style:key name="x">0.0</style:key>       <style:key name="y">45.0</style:key>
        <style:key name="width">110.0</style:key> <style:key name="height">15.0</style:key>
        <style:key name="font-name">Arial</style:key> <style:key name="font-size">9</style:key>
        <style:key name="font-bold">true</style:key>  <style:key name="text-color">#2D3748</style:key>
      </style:element-style>
      <core:value>Aanvoer / Afvoer:</core:value>
    </label>
    <message>
      <style:element-style>
        <style:key name="x">115.0</style:key>     <style:key name="y">45.0</style:key>
        <style:key name="width">408.0</style:key> <style:key name="height">15.0</style:key>
        <style:key name="font-name">Arial</style:key> <style:key name="font-size">9</style:key>
        <style:key name="font-bold">true</style:key>  <style:key name="text-color">#1E3A8A</style:key>
      </style:element-style>
      <core:value>$(AfvaltransportAanvoer__afvoer)</core:value>
    </message>

    <!-- ======================================================= -->
    <!-- SECTION 1 – CHAUFFEUR &amp; VOERTUIG                    -->
    <!-- ======================================================= -->
    <rectangle>
      <style:element-style>
        <style:key name="x">0.0</style:key>       <style:key name="y">65.0</style:key>
        <style:key name="width">523.0</style:key> <style:key name="height">35.0</style:key>
        <style:key name="stroke-color">#CBD5E0</style:key> <style:key name="stroke-width">1.0</style:key>
        <style:key name="draw-shape">true</style:key> <style:key name="fill-shape">false</style:key>
      </style:element-style>
    </rectangle>
    <rectangle>
      <style:element-style>
        <style:key name="x">0.0</style:key>       <style:key name="y">65.0</style:key>
        <style:key name="width">523.0</style:key> <style:key name="height">15.0</style:key>
        <style:key name="fill-color">#EDF2F7</style:key>
        <style:key name="fill-shape">true</style:key> <style:key name="draw-shape">false</style:key>
      </style:element-style>
    </rectangle>
    <label>
      <style:element-style>
        <style:key name="x">5.0</style:key> <style:key name="y">67.0</style:key>
        <style:key name="width">513.0</style:key> <style:key name="height">12.0</style:key>
        <style:key name="font-name">Arial</style:key> <style:key name="font-size">8</style:key>
        <style:key name="font-bold">true</style:key>  <style:key name="text-color">#4A5568</style:key>
      </style:element-style>
      <core:value>CHAUFFEUR &amp; VOERTUIG</core:value>
    </label>
    <label>
      <style:element-style>
        <style:key name="x">5.0</style:key>  <style:key name="y">83.0</style:key>
        <style:key name="width">60.0</style:key> <style:key name="height">12.0</style:key>
        <style:key name="font-name">Arial</style:key> <style:key name="font-size">8</style:key>
        <style:key name="font-bold">true</style:key>  <style:key name="text-color">#4A5568</style:key>
      </style:element-style>
      <core:value>Chauffeur:</core:value>
    </label>
    <message>
      <style:element-style>
        <style:key name="x">70.0</style:key> <style:key name="y">83.0</style:key>
        <style:key name="width">180.0</style:key> <style:key name="height">12.0</style:key>
        <style:key name="font-name">Arial</style:key> <style:key name="font-size">8</style:key>
        <style:key name="text-color">#1E3A8A</style:key>
      </style:element-style>
      <core:value>$(WerknemerNaam) $(WerknemerVoornaam)</core:value>
    </message>
    <label>
      <style:element-style>
        <style:key name="x">270.0</style:key> <style:key name="y">83.0</style:key>
        <style:key name="width">70.0</style:key>  <style:key name="height">12.0</style:key>
        <style:key name="font-name">Arial</style:key> <style:key name="font-size">8</style:key>
        <style:key name="font-bold">true</style:key>  <style:key name="text-color">#4A5568</style:key>
      </style:element-style>
      <core:value>Nummerplaat:</core:value>
    </label>
    <message>
      <style:element-style>
        <style:key name="x">345.0</style:key> <style:key name="y">83.0</style:key>
        <style:key name="width">170.0</style:key> <style:key name="height">12.0</style:key>
        <style:key name="font-name">Arial</style:key> <style:key name="font-size">8</style:key>
        <style:key name="text-color">#1E3A8A</style:key>
      </style:element-style>
      <core:value>$(VoertuigenNummerplaat)</core:value>
    </message>

    <!-- ======================================================= -->
    <!-- SECTION 2 – OPDRACHTGEVER | OVERBRENGER                 -->
    <!-- ======================================================= -->
    <rectangle>
      <style:element-style>
        <style:key name="x">0.0</style:key>       <style:key name="y">110.0</style:key>
        <style:key name="width">523.0</style:key> <style:key name="height">55.0</style:key>
        <style:key name="stroke-color">#CBD5E0</style:key> <style:key name="stroke-width">1.0</style:key>
        <style:key name="draw-shape">true</style:key> <style:key name="fill-shape">false</style:key>
      </style:element-style>
    </rectangle>
    <rectangle>
      <style:element-style>
        <style:key name="x">0.0</style:key>       <style:key name="y">110.0</style:key>
        <style:key name="width">523.0</style:key> <style:key name="height">15.0</style:key>
        <style:key name="fill-color">#EDF2F7</style:key>
        <style:key name="fill-shape">true</style:key> <style:key name="draw-shape">false</style:key>
      </style:element-style>
    </rectangle>
    <rectangle>
      <style:element-style>
        <style:key name="x">261.0</style:key> <style:key name="y">110.0</style:key>
        <style:key name="width">1.0</style:key>   <style:key name="height">55.0</style:key>
        <style:key name="fill-color">#CBD5E0</style:key>
        <style:key name="fill-shape">true</style:key> <style:key name="draw-shape">false</style:key>
      </style:element-style>
    </rectangle>
    <label>
      <style:element-style>
        <style:key name="x">5.0</style:key>   <style:key name="y">112.0</style:key>
        <style:key name="width">250.0</style:key> <style:key name="height">12.0</style:key>
        <style:key name="font-name">Arial</style:key> <style:key name="font-size">8</style:key>
        <style:key name="font-bold">true</style:key>  <style:key name="text-color">#4A5568</style:key>
      </style:element-style>
      <core:value>OPDRACHTGEVER</core:value>
    </label>
    <label>
      <style:element-style>
        <style:key name="x">270.0</style:key> <style:key name="y">112.0</style:key>
        <style:key name="width">250.0</style:key> <style:key name="height">12.0</style:key>
        <style:key name="font-name">Arial</style:key> <style:key name="font-size">8</style:key>
        <style:key name="font-bold">true</style:key>  <style:key name="text-color">#4A5568</style:key>
      </style:element-style>
      <core:value>OVERBRENGER</core:value>
    </label>
    <label>
      <style:element-style>
        <style:key name="x">5.0</style:key>  <style:key name="y">128.0</style:key>
        <style:key name="width">40.0</style:key> <style:key name="height">12.0</style:key>
        <style:key name="font-name">Arial</style:key> <style:key name="font-size">8</style:key>
        <style:key name="font-bold">true</style:key>  <style:key name="text-color">#4A5568</style:key>
      </style:element-style>
      <core:value>Naam:</core:value>
    </label>
    <message>
      <style:element-style>
        <style:key name="x">50.0</style:key> <style:key name="y">128.0</style:key>
        <style:key name="width">205.0</style:key> <style:key name="height">12.0</style:key>
        <style:key name="font-name">Arial</style:key> <style:key name="font-size">8</style:key>
        <style:key name="text-color">#1E3A8A</style:key>
      </style:element-style>
      <core:value>$(WerkficheCode_1)</core:value>
    </message>
    <label>
      <style:element-style>
        <style:key name="x">5.0</style:key>  <style:key name="y">143.0</style:key>
        <style:key name="width">40.0</style:key> <style:key name="height">12.0</style:key>
        <style:key name="font-name">Arial</style:key> <style:key name="font-size">8</style:key>
        <style:key name="font-bold">true</style:key>  <style:key name="text-color">#4A5568</style:key>
      </style:element-style>
      <core:value>Adres:</core:value>
    </label>
    <label>
      <style:element-style>
        <style:key name="x">270.0</style:key> <style:key name="y">128.0</style:key>
        <style:key name="width">40.0</style:key> <style:key name="height">12.0</style:key>
        <style:key name="font-name">Arial</style:key> <style:key name="font-size">8</style:key>
        <style:key name="font-bold">true</style:key>  <style:key name="text-color">#4A5568</style:key>
      </style:element-style>
      <core:value>Naam:</core:value>
    </label>
    <message>
      <style:element-style>
        <style:key name="x">315.0</style:key> <style:key name="y">128.0</style:key>
        <style:key name="width">200.0</style:key> <style:key name="height">12.0</style:key>
        <style:key name="font-name">Arial</style:key> <style:key name="font-size">8</style:key>
        <style:key name="text-color">#1E3A8A</style:key>
      </style:element-style>
      <core:value>$(WerkficheCode_1)</core:value>
    </message>
    <label>
      <style:element-style>
        <style:key name="x">270.0</style:key> <style:key name="y">143.0</style:key>
        <style:key name="width">40.0</style:key> <style:key name="height">12.0</style:key>
        <style:key name="font-name">Arial</style:key> <style:key name="font-size">8</style:key>
        <style:key name="font-bold">true</style:key>  <style:key name="text-color">#4A5568</style:key>
      </style:element-style>
      <core:value>Adres:</core:value>
    </label>

    <!-- ======================================================= -->
    <!-- SECTION 3 – HERKOMST | BESTEMMING                       -->
    <!-- ======================================================= -->
    <rectangle>
      <style:element-style>
        <style:key name="x">0.0</style:key>       <style:key name="y">175.0</style:key>
        <style:key name="width">523.0</style:key> <style:key name="height">85.0</style:key>
        <style:key name="stroke-color">#CBD5E0</style:key> <style:key name="stroke-width">1.0</style:key>
        <style:key name="draw-shape">true</style:key> <style:key name="fill-shape">false</style:key>
      </style:element-style>
    </rectangle>
    <rectangle>
      <style:element-style>
        <style:key name="x">0.0</style:key>       <style:key name="y">175.0</style:key>
        <style:key name="width">523.0</style:key> <style:key name="height">15.0</style:key>
        <style:key name="fill-color">#EDF2F7</style:key>
        <style:key name="fill-shape">true</style:key> <style:key name="draw-shape">false</style:key>
      </style:element-style>
    </rectangle>
    <rectangle>
      <style:element-style>
        <style:key name="x">261.0</style:key> <style:key name="y">175.0</style:key>
        <style:key name="width">1.0</style:key>   <style:key name="height">85.0</style:key>
        <style:key name="fill-color">#CBD5E0</style:key>
        <style:key name="fill-shape">true</style:key> <style:key name="draw-shape">false</style:key>
      </style:element-style>
    </rectangle>
    <label>
      <style:element-style>
        <style:key name="x">5.0</style:key>   <style:key name="y">177.0</style:key>
        <style:key name="width">250.0</style:key> <style:key name="height">12.0</style:key>
        <style:key name="font-name">Arial</style:key> <style:key name="font-size">8</style:key>
        <style:key name="font-bold">true</style:key>  <style:key name="text-color">#4A5568</style:key>
      </style:element-style>
      <core:value>HERKOMST</core:value>
    </label>
    <label>
      <style:element-style>
        <style:key name="x">270.0</style:key> <style:key name="y">177.0</style:key>
        <style:key name="width">250.0</style:key> <style:key name="height">12.0</style:key>
        <style:key name="font-name">Arial</style:key> <style:key name="font-size">8</style:key>
        <style:key name="font-bold">true</style:key>  <style:key name="text-color">#4A5568</style:key>
      </style:element-style>
      <core:value>BESTEMMING</core:value>
    </label>
    <label>
      <style:element-style>
        <style:key name="x">5.0</style:key>  <style:key name="y">193.0</style:key>
        <style:key name="width">40.0</style:key> <style:key name="height">12.0</style:key>
        <style:key name="font-name">Arial</style:key> <style:key name="font-size">8</style:key>
        <style:key name="font-bold">true</style:key>  <style:key name="text-color">#4A5568</style:key>
      </style:element-style>
      <core:value>Naam:</core:value>
    </label>
    <message>
      <style:element-style>
        <style:key name="x">50.0</style:key> <style:key name="y">193.0</style:key>
        <style:key name="width">205.0</style:key> <style:key name="height">12.0</style:key>
        <style:key name="font-name">Arial</style:key> <style:key name="font-size">8</style:key>
        <style:key name="text-color">#1E3A8A</style:key>
      </style:element-style>
      <core:value>$(WerfWerfnaam)</core:value>
    </message>
    <label>
      <style:element-style>
        <style:key name="x">5.0</style:key>  <style:key name="y">208.0</style:key>
        <style:key name="width">40.0</style:key> <style:key name="height">12.0</style:key>
        <style:key name="font-name">Arial</style:key> <style:key name="font-size">8</style:key>
        <style:key name="font-bold">true</style:key>  <style:key name="text-color">#4A5568</style:key>
      </style:element-style>
      <core:value>Adres:</core:value>
    </label>
    <message>
      <style:element-style>
        <style:key name="x">50.0</style:key> <style:key name="y">208.0</style:key>
        <style:key name="width">205.0</style:key> <style:key name="height">12.0</style:key>
        <style:key name="font-name">Arial</style:key> <style:key name="font-size">8</style:key>
        <style:key name="text-color">#1E3A8A</style:key>
      </style:element-style>
      <core:value>$(WerfWerfadresAdres_1)</core:value>
    </message>
    <message>
      <style:element-style>
        <style:key name="x">50.0</style:key> <style:key name="y">223.0</style:key>
        <style:key name="width">205.0</style:key> <style:key name="height">12.0</style:key>
        <style:key name="font-name">Arial</style:key> <style:key name="font-size">8</style:key>
        <style:key name="text-color">#1E3A8A</style:key>
      </style:element-style>
      <core:value>$(WerfWerfadresPostcode) $(WerfWerfadresGemeente)</core:value>
    </message>
    <message>
      <style:element-style>
        <style:key name="x">5.0</style:key> <style:key name="y">241.0</style:key>
        <style:key name="width">250.0</style:key> <style:key name="height">12.0</style:key>
        <style:key name="font-name">Arial</style:key> <style:key name="font-size">7</style:key>
        <style:key name="text-color">#1E3A8A</style:key>
      </style:element-style>
      <core:value>Werf nr: $(WerfWerfnr)   |   Project nr: $(WerfProjectProjectnr)</core:value>
    </message>
    <label>
      <style:element-style>
        <style:key name="x">270.0</style:key> <style:key name="y">193.0</style:key>
        <style:key name="width">40.0</style:key> <style:key name="height">12.0</style:key>
        <style:key name="font-name">Arial</style:key> <style:key name="font-size">8</style:key>
        <style:key name="font-bold">true</style:key>  <style:key name="text-color">#4A5568</style:key>
      </style:element-style>
      <core:value>Naam:</core:value>
    </label>
    <message>
      <style:element-style>
        <style:key name="x">315.0</style:key> <style:key name="y">193.0</style:key>
        <style:key name="width">200.0</style:key> <style:key name="height">12.0</style:key>
        <style:key name="font-name">Arial</style:key> <style:key name="font-size">8</style:key>
        <style:key name="text-color">#1E3A8A</style:key>
      </style:element-style>
      <core:value>$(VerwerkersBedrijfsnaam)</core:value>
    </message>
    <label>
      <style:element-style>
        <style:key name="x">270.0</style:key> <style:key name="y">208.0</style:key>
        <style:key name="width">40.0</style:key> <style:key name="height">12.0</style:key>
        <style:key name="font-name">Arial</style:key> <style:key name="font-size">8</style:key>
        <style:key name="font-bold">true</style:key>  <style:key name="text-color">#4A5568</style:key>
      </style:element-style>
      <core:value>Adres:</core:value>
    </label>
    <message>
      <style:element-style>
        <style:key name="x">315.0</style:key> <style:key name="y">208.0</style:key>
        <style:key name="width">200.0</style:key> <style:key name="height">12.0</style:key>
        <style:key name="font-name">Arial</style:key> <style:key name="font-size">8</style:key>
        <style:key name="text-color">#1E3A8A</style:key>
      </style:element-style>
      <core:value>$(VerwerkersStraat)</core:value>
    </message>
    <message>
      <style:element-style>
        <style:key name="x">315.0</style:key> <style:key name="y">223.0</style:key>
        <style:key name="width">200.0</style:key> <style:key name="height">12.0</style:key>
        <style:key name="font-name">Arial</style:key> <style:key name="font-size">8</style:key>
        <style:key name="text-color">#1E3A8A</style:key>
      </style:element-style>
      <core:value>$(VerwerkersPostcode) $(VerwerkersGemeente)</core:value>
    </message>

    <!-- ======================================================= -->
    <!-- SECTION 4 – AFVALOMSCHRIJVING                           -->
    <!-- ======================================================= -->
    <rectangle>
      <style:element-style>
        <style:key name="x">0.0</style:key>       <style:key name="y">270.0</style:key>
        <style:key name="width">523.0</style:key> <style:key name="height">65.0</style:key>
        <style:key name="stroke-color">#CBD5E0</style:key> <style:key name="stroke-width">1.0</style:key>
        <style:key name="draw-shape">true</style:key> <style:key name="fill-shape">false</style:key>
      </style:element-style>
    </rectangle>
    <rectangle>
      <style:element-style>
        <style:key name="x">0.0</style:key>       <style:key name="y">270.0</style:key>
        <style:key name="width">523.0</style:key> <style:key name="height">15.0</style:key>
        <style:key name="fill-color">#EDF2F7</style:key>
        <style:key name="fill-shape">true</style:key> <style:key name="draw-shape">false</style:key>
      </style:element-style>
    </rectangle>
    <label>
      <style:element-style>
        <style:key name="x">5.0</style:key> <style:key name="y">272.0</style:key>
        <style:key name="width">513.0</style:key> <style:key name="height">12.0</style:key>
        <style:key name="font-name">Arial</style:key> <style:key name="font-size">8</style:key>
        <style:key name="font-bold">true</style:key>  <style:key name="text-color">#4A5568</style:key>
      </style:element-style>
      <core:value>AFVALOMSCHRIJVING (Euralcode en omschrijving)</core:value>
    </label>
    <message>
      <style:element-style>
        <style:key name="x">5.0</style:key> <style:key name="y">291.0</style:key>
        <style:key name="width">513.0</style:key> <style:key name="height">15.0</style:key>
        <style:key name="font-name">Arial</style:key> <style:key name="font-size">9</style:key>
        <style:key name="text-color">#1E3A8A</style:key>
      </style:element-style>
      <core:value>$(AfvaltransportRecyclage)</core:value>
    </message>
    <message>
      <style:element-style>
        <style:key name="x">5.0</style:key> <style:key name="y">311.0</style:key>
        <style:key name="width">513.0</style:key> <style:key name="height">15.0</style:key>
        <style:key name="font-name">Arial</style:key> <style:key name="font-size">9</style:key>
        <style:key name="font-bold">true</style:key> <style:key name="text-color">#1E3A8A</style:key>
      </style:element-style>
      <core:value>Euralcode: $(Euralcode)</core:value>
    </message>

    <!-- ======================================================= -->
    <!-- SECTION 5 – OPMERKINGEN                                 -->
    <!-- ======================================================= -->
    <rectangle>
      <style:element-style>
        <style:key name="x">0.0</style:key>       <style:key name="y">345.0</style:key>
        <style:key name="width">523.0</style:key> <style:key name="height">55.0</style:key>
        <style:key name="stroke-color">#CBD5E0</style:key> <style:key name="stroke-width">1.0</style:key>
        <style:key name="draw-shape">true</style:key> <style:key name="fill-shape">false</style:key>
      </style:element-style>
    </rectangle>
    <rectangle>
      <style:element-style>
        <style:key name="x">0.0</style:key>       <style:key name="y">345.0</style:key>
        <style:key name="width">523.0</style:key> <style:key name="height">15.0</style:key>
        <style:key name="fill-color">#EDF2F7</style:key>
        <style:key name="fill-shape">true</style:key> <style:key name="draw-shape">false</style:key>
      </style:element-style>
    </rectangle>
    <label>
      <style:element-style>
        <style:key name="x">5.0</style:key> <style:key name="y">347.0</style:key>
        <style:key name="width">513.0</style:key> <style:key name="height">12.0</style:key>
        <style:key name="font-name">Arial</style:key> <style:key name="font-size">8</style:key>
        <style:key name="font-bold">true</style:key>  <style:key name="text-color">#4A5568</style:key>
      </style:element-style>
      <core:value>OPMERKINGEN TE VERMELDEN OP STORTBON</core:value>
    </label>
    <message>
      <style:element-style>
        <style:key name="x">5.0</style:key> <style:key name="y">368.0</style:key>
        <style:key name="width">513.0</style:key> <style:key name="height">25.0</style:key>
        <style:key name="font-name">Arial</style:key> <style:key name="font-size">9</style:key>
        <style:key name="text-color">#1E3A8A</style:key>
      </style:element-style>
      <core:value>Werkfiche nr: $(WerkficheNr)   |   $(WerkficheNaam)</core:value>
    </message>

    <!-- ======================================================= -->
    <!-- SECTION 6 – HANDTEKENINGEN                              -->
    <!-- ======================================================= -->
    <rectangle>
      <style:element-style>
        <style:key name="x">0.0</style:key>       <style:key name="y">410.0</style:key>
        <style:key name="width">523.0</style:key> <style:key name="height">140.0</style:key>
        <style:key name="stroke-color">#CBD5E0</style:key> <style:key name="stroke-width">1.0</style:key>
        <style:key name="draw-shape">true</style:key> <style:key name="fill-shape">false</style:key>
      </style:element-style>
    </rectangle>
    <rectangle>
      <style:element-style>
        <style:key name="x">0.0</style:key>       <style:key name="y">410.0</style:key>
        <style:key name="width">523.0</style:key> <style:key name="height">15.0</style:key>
        <style:key name="fill-color">#EDF2F7</style:key>
        <style:key name="fill-shape">true</style:key> <style:key name="draw-shape">false</style:key>
      </style:element-style>
    </rectangle>
    <!-- vertical separators -->
    <rectangle>
      <style:element-style>
        <style:key name="x">172.0</style:key> <style:key name="y">410.0</style:key>
        <style:key name="width">1.0</style:key> <style:key name="height">140.0</style:key>
        <style:key name="fill-color">#CBD5E0</style:key>
        <style:key name="fill-shape">true</style:key> <style:key name="draw-shape">false</style:key>
      </style:element-style>
    </rectangle>
    <rectangle>
      <style:element-style>
        <style:key name="x">347.0</style:key> <style:key name="y">410.0</style:key>
        <style:key name="width">1.0</style:key> <style:key name="height">140.0</style:key>
        <style:key name="fill-color">#CBD5E0</style:key>
        <style:key name="fill-shape">true</style:key> <style:key name="draw-shape">false</style:key>
      </style:element-style>
    </rectangle>
    <!-- column headers -->
    <label>
      <style:element-style>
        <style:key name="x">5.0</style:key>   <style:key name="y">412.0</style:key>
        <style:key name="width">160.0</style:key> <style:key name="height">12.0</style:key>
        <style:key name="font-name">Arial</style:key> <style:key name="font-size">7</style:key>
        <style:key name="font-bold">true</style:key>  <style:key name="text-color">#4A5568</style:key>
      </style:element-style>
      <core:value>OPDRACHTGEVER / PRODUCENT</core:value>
    </label>
    <label>
      <style:element-style>
        <style:key name="x">180.0</style:key> <style:key name="y">412.0</style:key>
        <style:key name="width">160.0</style:key> <style:key name="height">12.0</style:key>
        <style:key name="font-name">Arial</style:key> <style:key name="font-size">7</style:key>
        <style:key name="font-bold">true</style:key>  <style:key name="text-color">#4A5568</style:key>
      </style:element-style>
      <core:value>VERVOERDER</core:value>
    </label>
    <label>
      <style:element-style>
        <style:key name="x">355.0</style:key> <style:key name="y">412.0</style:key>
        <style:key name="width">160.0</style:key> <style:key name="height">12.0</style:key>
        <style:key name="font-name">Arial</style:key> <style:key name="font-size">7</style:key>
        <style:key name="font-bold">true</style:key>  <style:key name="text-color">#4A5568</style:key>
      </style:element-style>
      <core:value>ONTVANGER</core:value>
    </label>
    <!-- sub-labels -->
    <label>
      <style:element-style>
        <style:key name="x">5.0</style:key>   <style:key name="y">430.0</style:key>
        <style:key name="width">160.0</style:key> <style:key name="height">12.0</style:key>
        <style:key name="font-name">Arial</style:key> <style:key name="font-size">8</style:key>
        <style:key name="font-italic">true</style:key> <style:key name="text-color">#718096</style:key>
      </style:element-style>
      <core:value>Naam + datum + handtekening</core:value>
    </label>
    <label>
      <style:element-style>
        <style:key name="x">180.0</style:key> <style:key name="y">430.0</style:key>
        <style:key name="width">160.0</style:key> <style:key name="height">12.0</style:key>
        <style:key name="font-name">Arial</style:key> <style:key name="font-size">8</style:key>
        <style:key name="font-italic">true</style:key> <style:key name="text-color">#718096</style:key>
      </style:element-style>
      <core:value>Naam + datum + handtekening</core:value>
    </label>
    <message>
      <style:element-style>
        <style:key name="x">180.0</style:key> <style:key name="y">448.0</style:key>
        <style:key name="width">160.0</style:key> <style:key name="height">12.0</style:key>
        <style:key name="font-name">Arial</style:key> <style:key name="font-size">8</style:key>
        <style:key name="text-color">#1E3A8A</style:key>
      </style:element-style>
      <core:value>$(WerkficheCode_1)</core:value>
    </message>
    <message>
      <style:element-style>
        <style:key name="x">180.0</style:key> <style:key name="y">463.0</style:key>
        <style:key name="width">160.0</style:key> <style:key name="height">12.0</style:key>
        <style:key name="font-name">Arial</style:key> <style:key name="font-size">8</style:key>
        <style:key name="text-color">#1E3A8A</style:key>
      </style:element-style>
      <core:value>Datum: $(EventStart_Day)</core:value>
    </message>
    <message>
      <style:element-style>
        <style:key name="x">180.0</style:key> <style:key name="y">478.0</style:key>
        <style:key name="width">160.0</style:key> <style:key name="height">12.0</style:key>
        <style:key name="font-name">Arial</style:key> <style:key name="font-size">8</style:key>
        <style:key name="text-color">#1E3A8A</style:key>
      </style:element-style>
      <core:value>$(WerknemerNaam) $(WerknemerVoornaam)</core:value>
    </message>
    <label>
      <style:element-style>
        <style:key name="x">355.0</style:key> <style:key name="y">430.0</style:key>
        <style:key name="width">160.0</style:key> <style:key name="height">12.0</style:key>
        <style:key name="font-name">Arial</style:key> <style:key name="font-size">8</style:key>
        <style:key name="font-italic">true</style:key> <style:key name="text-color">#718096</style:key>
      </style:element-style>
      <core:value>Naam + datum + handtekening</core:value>
    </label>
    <message>
      <style:element-style>
        <style:key name="x">355.0</style:key> <style:key name="y">448.0</style:key>
        <style:key name="width">160.0</style:key> <style:key name="height">12.0</style:key>
        <style:key name="font-name">Arial</style:key> <style:key name="font-size">8</style:key>
        <style:key name="text-color">#1E3A8A</style:key>
      </style:element-style>
      <core:value>$(VerwerkersBedrijfsnaam)</core:value>
    </message>

  </report-header>

  <body>
    <details/>
  </body>

</layout>'
Add-ZipEntry $zip "layout.xml" $layout

# -----------------------------------------------------------------------
# Close the archive
# -----------------------------------------------------------------------
$zip.Dispose()
$fileStream.Close()

Write-Host "Done! Stortformulier.prpt is ready ($prptPath)"
