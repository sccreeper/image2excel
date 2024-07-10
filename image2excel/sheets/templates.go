package sheets

import _ "embed"

const WorkbookTemplateStart string = `<?xml version="1.0" encoding="UTF-8"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:ap="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:op="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:cdr="http://schemas.openxmlformats.org/drawingml/2006/chartDrawing" xmlns:comp="http://schemas.openxmlformats.org/drawingml/2006/compatibility" xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" xmlns:lc="http://schemas.openxmlformats.org/drawingml/2006/lockedCanvas" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:ds="http://schemas.openxmlformats.org/officeDocument/2006/customXml" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:sl="http://schemas.openxmlformats.org/schemaLibrary/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:xne="http://schemas.microsoft.com/office/excel/2006/main" xmlns:mso="http://schemas.microsoft.com/office/2006/01/customui" xmlns:ax="http://schemas.microsoft.com/office/2006/activeX" xmlns:cppr="http://schemas.microsoft.com/office/2006/coverPageProps" xmlns:cdip="http://schemas.microsoft.com/office/2006/customDocumentInformationPanel" xmlns:ct="http://schemas.microsoft.com/office/2006/metadata/contentType" xmlns:ntns="http://schemas.microsoft.com/office/2006/metadata/customXsn" xmlns:lp="http://schemas.microsoft.com/office/2006/metadata/longProperties" xmlns:ma="http://schemas.microsoft.com/office/2006/metadata/properties/metaAttributes" xmlns:msink="http://schemas.microsoft.com/ink/2010/main" xmlns:c14="http://schemas.microsoft.com/office/drawing/2007/8/2/chart" xmlns:cdr14="http://schemas.microsoft.com/office/drawing/2010/chartDrawing" xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main" xmlns:pic14="http://schemas.microsoft.com/office/drawing/2010/picture" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" xmlns:xdr14="http://schemas.microsoft.com/office/excel/2010/spreadsheetDrawing" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" xmlns:dsp="http://schemas.microsoft.com/office/drawing/2008/diagram" xmlns:mso14="http://schemas.microsoft.com/office/2009/07/customui" xmlns:dgm14="http://schemas.microsoft.com/office/drawing/2010/diagram" xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main" xmlns:x12ac="http://schemas.microsoft.com/office/spreadsheetml/2011/1/ac" xmlns:x15ac="http://schemas.microsoft.com/office/spreadsheetml/2010/11/ac" xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision" xmlns:xr2="http://schemas.microsoft.com/office/spreadsheetml/2015/revision2" xmlns:xr3="http://schemas.microsoft.com/office/spreadsheetml/2016/revision3" xmlns:xr4="http://schemas.microsoft.com/office/spreadsheetml/2016/revision4" xmlns:xr5="http://schemas.microsoft.com/office/spreadsheetml/2016/revision5" xmlns:xr6="http://schemas.microsoft.com/office/spreadsheetml/2016/revision6" xmlns:xr7="http://schemas.microsoft.com/office/spreadsheetml/2016/revision7" xmlns:xr8="http://schemas.microsoft.com/office/spreadsheetml/2016/revision8" xmlns:xr9="http://schemas.microsoft.com/office/spreadsheetml/2016/revision9" xmlns:xr10="http://schemas.microsoft.com/office/spreadsheetml/2016/revision10" xmlns:xr11="http://schemas.microsoft.com/office/spreadsheetml/2016/revision11" xmlns:xr12="http://schemas.microsoft.com/office/spreadsheetml/2016/revision12" xmlns:xr13="http://schemas.microsoft.com/office/spreadsheetml/2016/revision13" xmlns:xr14="http://schemas.microsoft.com/office/spreadsheetml/2016/revision14" xmlns:xr15="http://schemas.microsoft.com/office/spreadsheetml/2016/revision15" xmlns:x16="http://schemas.microsoft.com/office/spreadsheetml/2014/11/main" xmlns:x16r2="http://schemas.microsoft.com/office/spreadsheetml/2015/02/main" mc:Ignorable="c14 cdr14 a14 pic14 x14 xdr14 x14ac dsp mso14 dgm14 x15 x12ac x15ac xr xr2 xr3 xr4 xr5 xr6 xr7 xr8 xr9 xr10 xr11 xr12 xr13 xr14 xr15 x15 x16 x16r2 mo mx mv o v" xmlns:mo="http://schemas.microsoft.com/office/mac/office/2008/main" xmlns:mx="http://schemas.microsoft.com/office/mac/excel/2008/main" xmlns:mv="urn:schemas-microsoft-com:mac:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:v="urn:schemas-microsoft-com:vml" xr:uid="{00000000-0001-0000-0000-000000000000}">
<dimension ref="A1"/>
<sheetViews>
<sheetView tabSelected="true" workbookViewId="0"/>
</sheetViews>
<cols>
<col min="1" max="%d" width="2.5" customWidth="1"/>
</cols>
<sheetData>
`

const WorkbookTemplateEnd string = `</worksheet>`

const ConditionalFormattingRule string = `<conditionalFormatting sqref="%s:%s"><cfRule type="colorScale" priority="%d"><colorScale><cfvo type="min" val="0"/><cfvo type="max" val="0"/><color rgb="FF000000"/><color rgb="FF%s"/></colorScale></cfRule></conditionalFormatting>`

// /_rels/.rels
const RootRelationshipsTemplate string = `<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId3" Target="docProps/app.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties"></Relationship><Relationship Id="rId2" Target="docProps/core.xml" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties"></Relationship><Relationship Id="rId1" Target="xl/workbook.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"></Relationship></Relationships>`

// /docProps/app.xml
const DocPropsAppTemplate string = `<?xml version="1.0" encoding="UTF-8"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><TotalTime>0</TotalTime><Application>image2excel</Application></Properties>`

// /docProps/core.xml
const DocPropsCoreTemplate string = `<?xml version="1.0" encoding="UTF-8"?>
<cp:coreProperties
    xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
    xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/"
    xmlns:dcmitype="http://purl.org/dc/dcmitype/"
    xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
    <dc:creator>xuri</dc:creator>
    <dcterms:created xsi:type="dcterms:W3CDTF">2006-09-16T00:00:00Z</dcterms:created>
    <dcterms:modified xsi:type="dcterms:W3CDTF">2006-09-16T00:00:00Z</dcterms:modified>
</cp:coreProperties>`

const RelationshipIdStart int = 4

const XlRelationshipsStartTemplate string = `<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId2" Target="styles.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"></Relationship><Relationship Id="rId3" Target="theme/theme1.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"></Relationship>`

//go:embed large/theme1.xml
var ThemeOneXml []byte

//go:embed large/styles.xml
var StylesXml []byte

const WorkbookMetaStartXml string = `<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x15"
xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main">
<fileVersion appName="xl" lastEdited="6" lowestEdited="6" rupBuild="14420"></fileVersion>
<workbookPr filterPrivacy="true" defaultThemeVersion="164011"></workbookPr>
<bookViews>
<workbookView xWindow="0" yWindow="0" windowWidth="14805" windowHeight="8010" activeTab="14"></workbookView>
</bookViews>
<sheets>`

const WorkbookMetaEndXml string = `</sheets>
    <calcPr calcId="122211"></calcPr>
</workbook>`

const ContentTypesStartXml string = `<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
    <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"></Default>
    <Default Extension="xml" ContentType="application/xml"></Default>
    <Default Extension="bin" ContentType="application/vnd.ms-office.vbaProject"></Default>
    <Override PartName="/xl/theme/theme1.xml"
        ContentType="application/vnd.openxmlformats-officedocument.theme+xml"></Override>
    <Override PartName="/xl/styles.xml"
        ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"></Override>
    <Override PartName="/xl/workbook.xml"
        ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"></Override>
    <Override PartName="/docProps/app.xml"
        ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"></Override>
    <Override PartName="/docProps/core.xml"
        ContentType="application/vnd.openxmlformats-package.core-properties+xml"></Override>`

const contentTypesEndXml string = `</Types>`
