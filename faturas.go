package main

import(
 "fmt"
 "os"
 "io"
 "log"
 "time"
 "strings"
 "strconv"
 "archive/zip"
 "path/filepath"
 "github.com/antchfx/xmlquery"
 "github.com/xuri/excelize/v2"
)

type Data struct {
  Proveedor string
  NitProveedor string
  Fecha string
  Cliente string
  Base string
  Iva19 string
  Iva5 string
  ReteIca string
  Total string
}

func parseXML(path string) (Data, error) {
  var datos Data
  d, err := os.ReadFile(path)
  if err != nil {
    return datos, err
  }
  
  doc, err := xmlquery.Parse(strings.NewReader(string(d)))
  doc2, err := xmlquery.Parse(strings.NewReader(string(getText(doc, "//cbc:Description"))))
  if err != nil {
    return datos, err
  }
  
  datos = Data{
    Proveedor: getText(doc2, "//cac:PartyTaxScheme//cbc:RegistrationName"),
		NitProveedor: getText(doc2, "//cac:PartyTaxScheme//cbc:CompanyID"),
		Cliente: getText(doc2, "//cac:AccountingCustomerParty//cac:Party//cac:PartyName//cbc:Name"),
		Total: getText(doc2, "//cbc:PayableAmount"),
		Fecha: getText(doc2, "//cbc:IssueDate"),

		Base: getText(doc2, "//cac:LegalMonetaryTotal//cbc:LineExtensionAmount"),
		Iva19: getText(doc2, "//*[local-name()='TaxSubtotal'][.//*[local-name()='Percent']='19.00']//*[local-name()='TaxAmount']"),

		Iva5: getText(doc2, "//*[local-name()='TaxSubtotal'][.//*[local-name()='Percent']='5.00']//*[local-name()='TaxAmount']"),
		ReteIca: getText(doc2, "//*[local-name()='TaxScheme']/*[local-name()='Name'][text()='IC']/../../..//*[local-name()='TaxAmount']"),
  }
  
  return datos, nil
}

func getText(doc *xmlquery.Node, path string) string {
  node := xmlquery.FindOne(doc, path)
	if node != nil {
		return node.InnerText()
	}
	return ""
}

func procesarZip(path string) (Data, error) {
  var factura Data
  var err error
  
  r, err := zip.OpenReader(path)
  if err != nil {
      return factura, err
  }
  defer r.Close()

  var xmlPath, pdfPath string

  // Extraer archivos
  for _, file := range r.File {
    rc, err := file.Open()
    if err != nil {
        return factura, err
    }
    defer rc.Close()

    outPath := filepath.Join(".", file.Name)
    outFile, err := os.Create(outPath)
    if err != nil {
        return factura, err
    }

    _, err = io.Copy(outFile, rc)
    outFile.Close()
    if err != nil {
        return factura, err
    }

    if strings.HasSuffix(strings.ToLower(file.Name), ".xml") {
        xmlPath = outPath
    } else if strings.HasSuffix(strings.ToLower(file.Name), ".pdf") {
        pdfPath = outPath
    }
  }

  if xmlPath != "" && pdfPath != "" {
    factura, err = parseXML(xmlPath)
    if err != nil {
      return factura, err
    }

    destinoDir := "pdfs_" + factura.Proveedor
    os.MkdirAll(destinoDir, os.ModePerm)

    nuevoNombre := strings.ReplaceAll(factura.Proveedor  + "_" + factura.Fecha + "_" + factura.Total, " ", "_") + ".pdf"
    destino := filepath.Join(destinoDir, nuevoNombre)

    if err := os.Rename(pdfPath, destino); err != nil {
        return factura, err
    }
    
    os.Remove(xmlPath)

    //fmt.Println("PDF movido a:", destino)
  }
	
  return factura, err
}

func openExcel(filePath string, sheet string) (*excelize.File, int, error) {
	var f *excelize.File
	var err error

	if _, err := os.Stat(filePath); os.IsNotExist(err) {
		// Si no existe, crear con encabezados
		f = excelize.NewFile()
		f.SetSheetName("Sheet1", sheet)
		headers := []string{"Proveedor", "NIT Proveedor", "Fecha", "Cliente", "Base", "IVA 19%", "IVA 5%", "ReteICA", "Total"}
		for i, header := range headers {
			cell, _ := excelize.CoordinatesToCellName(i+1, 1)
			f.SetCellValue(sheet, cell, header)
		}
		return f, 2, nil // siguiente fila libre es la 2
	}

	// Si ya existe, abrir
	f, err = excelize.OpenFile(filePath)
	if err != nil {
		return nil, 0, err
	}
	rows, err := f.GetRows(sheet)
	if err != nil {
		return nil, 0, err
	}
	return f, len(rows) + 1, nil
}

func appendRow(f *excelize.File, sheet string, rowNum int, data Data) {
	// Datos en el orden deseado
	values := []string{
		data.Proveedor,
		data.NitProveedor,
		data.Fecha,
		data.Cliente,
		data.Base,
		data.Iva19,
		data.Iva5,
		data.ReteIca,
		data.Total,
	}

	// Formatos
	textStyle, _ := f.NewStyle(&excelize.Style{NumFmt: 49})  // Texto
	dateStyle, _ := f.NewStyle(&excelize.Style{NumFmt: 14})  // Fecha corta (dd/mm/yyyy)
	moneyStyle, _ := f.NewStyle(&excelize.Style{NumFmt: 4})  // Moneda con separador de miles
	numberStyle, _ := f.NewStyle(&excelize.Style{NumFmt: 2}) // Número decimal con 2 cifras

	// Formato esperado para fecha
	layout := "02/01/2006" // dd/mm/yyyy

	for i, val := range values {
		cell, _ := excelize.CoordinatesToCellName(i+1, rowNum)

		switch i {
		case 2: // Fecha
			if t, err := time.Parse(layout, val); err == nil {
				f.SetCellValue(sheet, cell, t)
				f.SetCellStyle(sheet, cell, cell, dateStyle)
				continue
			}

		case 4, 5, 6, 7, 8: // Montos y números
			if num, err := strconv.ParseFloat(strings.ReplaceAll(val, ",", ""), 64); err == nil {
				if i == 4 || i == 8 { // Base y Total → moneda
					f.SetCellValue(sheet, cell, num)
					f.SetCellStyle(sheet, cell, cell, moneyStyle)
				} else { // IVA y retenciones → número con decimales
					f.SetCellValue(sheet, cell, num)
					f.SetCellStyle(sheet, cell, cell, numberStyle)
				}
				continue
			}
		}

		// Por defecto, texto
		f.SetCellValue(sheet, cell, val)
		f.SetCellStyle(sheet, cell, cell, textStyle)
	}
}

func main() {
  
  carpeta := "."
  
  excelPath := "./facturas.xlsx"
  sheetName := "Facturas"
  
  var facturas []Data
  
  // Abrir o crear Excel una sola vez
	f, nextRow, err := openExcel(excelPath, sheetName)
	if err != nil {
		log.Fatal(err)
	}
	defer f.Close()

  filepath.Walk(carpeta, func(path string, info os.FileInfo, err error) error {
      if err != nil {
          return err
      }
      if !info.IsDir() && strings.HasSuffix(strings.ToLower(path), ".zip") {
          fmt.Println("Procesando:", path)
          factura, err := procesarZip(path)
          if err != nil {
              fmt.Println("Error procesando zip:", err)
          }
          facturas = append(facturas, factura)
          os.Remove(path)
      }
      return nil
  })
  
  for _, factura := range facturas {
    appendRow(f, sheetName, nextRow, factura)
    nextRow++
  }
  
  if err := f.SaveAs(excelPath); err != nil {
		log.Fatal(err)
	}
}