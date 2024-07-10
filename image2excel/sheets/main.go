package sheets

import (
	"archive/zip"
	"errors"
	"fmt"
	"image"
	"os"

	"github.com/xuri/excelize/v2"
)

// Package "sheets"
// Package used internally for written specifically writing to Excel spreadsheets for the purpose of converting images to spreadsheets.

type Workbook struct {
	Path   string
	Sheets []*Sheet
}

type Sheet struct {
	Name      string
	xmlString []byte
}

func NewWorkbook(path string) *Workbook {

	return &Workbook{Path: path, Sheets: make([]*Sheet, 0)}
}

func (wb *Workbook) AddSheet(name string) *Sheet {

	wb.Sheets = append(wb.Sheets, &Sheet{Name: name, xmlString: make([]byte, 0)})

	return wb.Sheets[len(wb.Sheets)-1]
}

func (sheet *Sheet) AddImage(image image.Image) (err error) {

	sheet.xmlString = []byte(fmt.Sprintf(WorkbookTemplateStart, image.Bounds().Dx()+2))

	// Add sheet data

	for y := 0; y < image.Bounds().Dy()*3; y += 2 {

		sheet.xmlString = append(sheet.xmlString, []byte(fmt.Sprintf(`<row r="%d">`, y+1))...)

		// Red

		for red := 0; red < image.Bounds().Dx(); red++ {

			r, _, _, _ := image.At(red, y/3).RGBA()

			cellCoords, err := excelize.CoordinatesToCellName(red+1, y+1)
			if err != nil {
				return err
			}

			sheet.xmlString = append(
				sheet.xmlString,
				[]byte(fmt.Sprintf(`<c r="%s"><v>%d</v></c>`, cellCoords, uint8(r/255)))...,
			)

		}

		endCoords1, err := excelize.CoordinatesToCellName(image.Bounds().Dx()+1, y+1)
		endCoords2, err1 := excelize.CoordinatesToCellName(image.Bounds().Dx()+2, y+1)
		if err != nil || err1 != nil {
			return errors.Join(err, err1)
		}

		sheet.xmlString = append(sheet.xmlString, []byte(fmt.Sprintf(`<c r="%s"><v>%d</v></c><c r="%s"><v>%d</v></c></row>`, endCoords1, 0, endCoords2, 255))...)
		sheet.xmlString = append(sheet.xmlString, []byte(fmt.Sprintf(`<row r="%d">`, y+2))...)
		// Green

		for green := 0; green < image.Bounds().Dx(); green++ {

			_, g, _, _ := image.At(green, y/3).RGBA()

			cellCoords, err := excelize.CoordinatesToCellName(green+1, y+2)
			if err != nil {
				return err
			}

			sheet.xmlString = append(
				sheet.xmlString,
				[]byte(fmt.Sprintf(`<c r="%s"><v>%d</v></c>`, cellCoords, uint8(g/255)))...,
			)

		}

		endCoords1, err = excelize.CoordinatesToCellName(image.Bounds().Dx()+1, y+2)
		endCoords2, err1 = excelize.CoordinatesToCellName(image.Bounds().Dx()+2, y+2)
		if err != nil || err1 != nil {
			return errors.Join(err, err1)
		}

		sheet.xmlString = append(sheet.xmlString, []byte(fmt.Sprintf(`<c r="%s"><v>%d</v></c><c r="%s"><v>%d</v></c></row>`, endCoords1, 0, endCoords2, 255))...)
		sheet.xmlString = append(sheet.xmlString, []byte(fmt.Sprintf(`<row r="%d">`, y+3))...)

		// Blue

		for blue := 0; blue < image.Bounds().Dx(); blue++ {

			_, _, b, _ := image.At(blue, y/3).RGBA()

			cellCoords, err := excelize.CoordinatesToCellName(blue+1, y+3)
			if err != nil {
				return err
			}

			sheet.xmlString = append(
				sheet.xmlString,
				[]byte(fmt.Sprintf(`<c r="%s"><v>%d</v></c>`, cellCoords, uint8(b/255)))...,
			)

		}

		endCoords1, err = excelize.CoordinatesToCellName(image.Bounds().Dx()+1, y+3)
		endCoords2, err1 = excelize.CoordinatesToCellName(image.Bounds().Dx()+2, y+3)
		if err != nil || err1 != nil {
			return errors.Join(err, err1)
		}

		sheet.xmlString = append(sheet.xmlString, []byte(fmt.Sprintf(`<c r="%s"><v>%d</v></c><c r="%s"><v>%d</v></c></row>`, endCoords1, 0, endCoords2, 255))...)
	}

	sheet.xmlString = append(sheet.xmlString, []byte("</sheetData>")...)

	// Add conditional formatting

	for y := 0; y < image.Bounds().Dy()*3; y += 3 {

		startCoords, err1 := excelize.CoordinatesToCellName(1, y+1)
		endCoords, err2 := excelize.CoordinatesToCellName(image.Bounds().Dx()+2, y+1)
		if err1 != nil || err2 != nil {
			return errors.Join(err1, err2)
		}

		sheet.xmlString = append(sheet.xmlString, []byte(fmt.Sprintf(ConditionalFormattingRule, startCoords, endCoords, y+1, "FF0000"))...)

		startCoords, err1 = excelize.CoordinatesToCellName(1, y+2)
		endCoords, err2 = excelize.CoordinatesToCellName(image.Bounds().Dx()+2, y+2)
		if err1 != nil || err2 != nil {
			return errors.Join(err1, err2)
		}

		sheet.xmlString = append(sheet.xmlString, []byte(fmt.Sprintf(ConditionalFormattingRule, startCoords, endCoords, y+1, "00FF00"))...)

		startCoords, err1 = excelize.CoordinatesToCellName(1, y+3)
		endCoords, err2 = excelize.CoordinatesToCellName(image.Bounds().Dx()+2, y+3)
		if err1 != nil || err2 != nil {
			return errors.Join(err1, err2)
		}

		sheet.xmlString = append(sheet.xmlString, []byte(fmt.Sprintf(ConditionalFormattingRule, startCoords, endCoords, y+1, "0000FF"))...)

	}

	sheet.xmlString = append(sheet.xmlString, []byte(WorkbookTemplateEnd)...)

	return nil

}

// Save the workbook, creating required files etc.
func (wb *Workbook) Save() (err error) {

	fileDisk, err := os.Create(wb.Path)
	if err != nil {
		return err
	}

	w := zip.NewWriter(fileDisk)

	// Create initial files

	rootRelationships, err := w.Create("_rels/.rels")
	if err != nil {
		return err
	}
	rootRelationships.Write([]byte(RootRelationshipsTemplate))

	docPropsApp, err := w.Create("docProps/app.xml")
	if err != nil {
		return err
	}
	docPropsApp.Write([]byte(DocPropsAppTemplate))

	docPropsCore, err := w.Create("docProps/core.xml")
	if err != nil {
		return err
	}
	docPropsCore.Write([]byte(DocPropsCoreTemplate))

	xlRelsWriter, err := w.Create("xl/_rels/workbook.xml.rels")
	if err != nil {
		return err
	}

	xlRelsString := XlRelationshipsStartTemplate

	for i := range wb.Sheets {

		xlRelsString += fmt.Sprintf(`<Relationship Id="rId%d" Target="/xl/worksheets/sheet%d.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"></Relationship>`, i+RelationshipIdStart, i+1)

	}

	xlRelsWriter.Write([]byte(xlRelsString))
	xlRelsWriter.Write([]byte("</Relationships>"))

	themeOneWriter, err := w.Create("xl/theme/theme1.xml")
	if err != nil {
		return err
	}
	themeOneWriter.Write(ThemeOneXml)

	stylesWriter, err := w.Create("xl/styles.xml")
	if err != nil {
		return err
	}
	stylesWriter.Write(StylesXml)

	// Create workbook XML file

	workbookMetaWriter, err := w.Create("xl/workbook.xml")
	if err != nil {
		return err
	}
	workbookMetaWriter.Write([]byte(WorkbookMetaStartXml))

	// Write sheet information

	for i, v := range wb.Sheets {
		workbookMetaWriter.Write([]byte(fmt.Sprintf(`<sheet name="%s" sheetId="%d" r:id="rId%d"></sheet>`, v.Name, i+1, i+4)))
	}
	workbookMetaWriter.Write([]byte(WorkbookMetaEndXml))

	// Write [Content_Types.xml]

	contentTypesWriter, err := w.Create("[Content_types.xml]")
	if err != nil {
		return err
	}

	contentTypesWriter.Write([]byte(ContentTypesStartXml))

	for i := range wb.Sheets {
		contentTypesWriter.Write([]byte(fmt.Sprintf(`<Override PartName="/xl/worksheets/sheet%d.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"></Override>`, i+1)))
	}

	contentTypesWriter.Write([]byte(contentTypesEndXml))

	// Finally write workbook XML

	for i, v := range wb.Sheets {

		sheetWriter, err := w.Create(fmt.Sprintf("xl/worksheets/sheet%d.xml", i+1))
		if err != nil {
			return err
		}
		sheetWriter.Write(v.xmlString)

	}

	w.Close()

	return

}
