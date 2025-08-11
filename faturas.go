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
  NumFactura string
  Proveedor string
  NitProveedor string
  Fecha string
  Cliente string
  Base string
  Descuento string
  Iva19 string
  Iva5 string
  Ico string
  ReteFuente string
  ReteIVA string
  ReteIca string
  Total string
  InfRt string
  Items string
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
  
  nodes := xmlquery.Find(doc2, "//cbc:Description")
  var descripciones []string
	for _, n := range nodes {
		descripciones = append(descripciones, n.InnerText())
	}
	// Unir todas las descripciones con punto y coma
	items := strings.Join(descripciones, "; ")
  
  taxNode := getText(doc2, "//cac:AccountingSupplierParty//cac:Party//cac:PartyTaxScheme//cbc:TaxLevelCode")
  taxNode = debeRetener(taxNode)
  
  datos = Data{
    NumFactura: getText(doc2, "//cbc:ID[not(ancestor::cac:*)]"),
    
    Proveedor: getText(doc2, "//cac:AccountingSupplierParty/cac:Party/cac:PartyTaxScheme/cbc:RegistrationName"),
    
		NitProveedor: getText(doc2, "//cac:AccountingSupplierParty/cac:Party/cac:PartyTaxScheme/cbc:CompanyID"),
	
		Cliente: getText(doc2, "//cac:AccountingCustomerParty/cac:Party/cac:PartyTaxScheme/cbc:RegistrationName"),
		
		Total: getText(doc2, "//cbc:PayableAmount"),
		
		Fecha: getText(doc2, "//cbc:IssueDate"),

		Base: getText(doc2, "//cac:LegalMonetaryTotal//cbc:LineExtensionAmount"),
		
		Descuento: getText(doc2, "//cac:LegalMonetaryTotal/cbc:AllowanceTotalAmount"),
		
		Iva19: getText(doc2, "//*[local-name()='TaxSubtotal'][.//*[local-name()='Percent']='19.00']//*[local-name()='TaxAmount']"),

		Iva5: getText(doc2, "//*[local-name()='TaxSubtotal'][.//*[local-name()='Percent']='5.00']//*[local-name()='TaxAmount']"),
		
		ReteIca: getText(doc2, "//*[local-name()='TaxScheme']/*[local-name()='ID'][text()='05']/../../..//*[local-name()='TaxAmount']"),
		
		ReteFuente: getText(doc2, "//*[local-name()='TaxScheme']/*[local-name()='ID'][text()='06']/../../..//*[local-name()='TaxAmount']"),
		
		ReteIVA: getText(doc2, "//*[local-name()='TaxScheme']/*[local-name()='ID'][text()='04']/../../..//*[local-name()='TaxAmount']"),
		
		Ico: getText(doc2, "//*[local-name()='TaxScheme']/*[local-name()='ID'][text()='02']/../../..//*[local-name()='TaxAmount']"),
		
		InfRt: taxNode,
		
		Items: items,
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
    destinoXML := "xmlDir"
    os.MkdirAll(destinoXML, os.ModePerm)
    

    nuevoNombre := strings.ReplaceAll(factura.Proveedor  + "_" + factura.NumFactura, " ", "_") + ".pdf"
    destino := filepath.Join(destinoDir, nuevoNombre)

    if err := os.Rename(pdfPath, destino); err != nil {
        return factura, err
    }
    
    nuevoNombreXML := strings.ReplaceAll(factura.Proveedor  + "_" + factura.NumFactura, " ", "_") + ".xml"
    
    destino = filepath.Join(destinoXML, nuevoNombreXML)
    
    if err := os.Rename(xmlPath, destino); err != nil {
        return factura, err
    }
    
    //os.Remove(xmlPath)

    //fmt.Println("PDF movido a:", destino)
  }
	
  return factura, err
}

func openExcel(filePath string, sheet string) (*excelize.File, [][]string, error) {
	var f *excelize.File
	var err error

	if _, err := os.Stat(filePath); os.IsNotExist(err) {
		// Si no existe, crear con encabezados
		f = excelize.NewFile()
		f.SetSheetName("Sheet1", sheet)
		headers := []string{"Factura", "Proveedor", "NIT Proveedor", "Fecha", "Cliente", "Base", "Descuento", "IVA 19%", "IVA 5%", "ICO", "Retefuente", "Reteiva", "Reteica", "Total", "Info Rete", "Items"}
		for i, header := range headers {
			cell, _ := excelize.CoordinatesToCellName(i+1, 1)
			f.SetCellValue(sheet, cell, header)
		}
		return f, [][]string{headers}, nil // siguiente fila libre es la 2
	}

	// Si ya existe, abrir
	f, err = excelize.OpenFile(filePath)
	if err != nil {
		return nil, [][]string{}, err
	}
	rows, err := f.GetRows(sheet)
	if err != nil {
		return nil, [][]string{}, err
	}
	
	return f, rows, nil
}

func appendRow(f *excelize.File, sheet string, rowNum int, data Data) {
	// Datos en el orden deseado
	values := []string{
	  data.NumFactura,
		data.Proveedor,
		data.NitProveedor,
		data.Fecha,
		data.Cliente,
		data.Base,
		data.Descuento,
		data.Iva19,
		data.Iva5,
		data.Ico,
		data.ReteFuente,
		data.ReteIVA,
		data.ReteIca,
		data.Total,
		data.InfRt,
		data.Items,
	}

	// Formatos
	textStyle, _ := f.NewStyle(&excelize.Style{NumFmt: 49})  // Texto
	dateStyle, _ := f.NewStyle(&excelize.Style{NumFmt: 14})  // Fecha corta (dd/mm/yyyy)
	moneyStyle, _ := f.NewStyle(&excelize.Style{NumFmt: 4})  // Moneda con separador de miles
	numberStyle, _ := f.NewStyle(&excelize.Style{NumFmt: 2}) // Número decimal con 2 cifras

	// Formato esperado para fecha
	layout := "2006/01/02" // yyyy/mm/dd

	for i, val := range values {
		cell, _ := excelize.CoordinatesToCellName(i+1, rowNum)

		switch i {
		case 3: // Fecha
		  val = strings.ReplaceAll(val, "-", "/")
			if t, err := time.Parse(layout, val); err == nil {
				f.SetCellValue(sheet, cell, t)
				f.SetCellStyle(sheet, cell, cell, dateStyle)
				continue
			}

		case 5, 6, 7, 8, 9, 10, 11, 12, 13: // Montos y números
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

// Mapa de códigos que NO retienen
var noRetencionCodes = map[string]bool{
    "O-15":    true, // Autorretenedor
    "O-47":    true, // Régimen simple de tributación
    "R-99-PN": true, // No responsable
    "O-49":    true, // No responsable de IVA
    "A-62":    true, // Importador Ocasional
    "A-64":    true, // Beneficiario Programa de Fomento Industria Automotriz-PROFIA
}

func debeRetener(codigos string) string {
  partes := strings.Split(codigos, ";")
  
  for _, c := range partes {
    c = strings.TrimSpace(c)
    // Si alguno NO está en la lista de no retención, entonces debe retener
    if noRetencionCodes[c] {
      // Si está en lista de no retención, no retiene
      return "No retención"
    }
  }
  
  return "Posible retención"
}

func main() {
  
  carpeta := "."
  
  var facturas []Data
  // Mover ZIP a carpeta zipsDir
  destinoDir := "zipsDir"
  os.MkdirAll(destinoDir, os.ModePerm)

  filepath.Walk(carpeta, func(path string, info os.FileInfo, err error) error {
    
    if err != nil {
        return err
    }
    
    // Saltar la carpeta zipsDir y su contenido
    if info.IsDir() && filepath.Base(path) == "zipsDir" {
      return filepath.SkipDir
    }
    
    if !info.IsDir() && strings.HasSuffix(strings.ToLower(path), ".zip") {
      //fmt.Println("Procesando:", path)
      factura, err := procesarZip(path)
      if err != nil {
          fmt.Println("Error procesando zip:", err)
      }
      facturas = append(facturas, factura)
      
      //os.Remove(path)
      nuevoNombreZip := strings.ReplaceAll(factura.Proveedor  + "_" + factura.NumFactura, " ", "_") + ".zip"
      
      destinoZip := filepath.Join(destinoDir, nuevoNombreZip)

      if err := os.Rename(path, destinoZip); err != nil {
        fmt.Println("Error moviendo zip:", err)
      }
    }
    return nil
  })
  
  excelPath := "./facturas.xlsx"
  sheetName := "Facturas"
  // Abrir o crear Excel una sola vez
	f, rows, err := openExcel(excelPath, sheetName)
	if err != nil {
		log.Fatal(err)
	}
	defer f.Close()
	
	m := make(map[string]bool, len(rows))
  for i := range rows {
    if len(rows[i]) < 2 {
        continue
    }
    clave := rows[i][0] + "|" + rows[i][1]
    m[clave] = true
  }
  
  nextRow := len(rows) + 1
  
  for _, factura := range facturas {
    claveNueva := factura.NumFactura + "|" + factura.Proveedor
    if m[claveNueva] {
        continue
    }
    
    appendRow(f, sheetName, nextRow, factura)
    nextRow++
  }
  
  if err := f.SaveAs(excelPath); err != nil {
		log.Fatal(err)
	}
}