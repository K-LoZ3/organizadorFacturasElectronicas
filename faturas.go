package main

import(
 "fmt"
 "os"
 "io"
 "strings"
 "archive/zip"
 "path/filepath"
 "github.com/antchfx/xmlquery"
 //"github.com/nwaples/rardecode"
 //"github.com/xuri/excelize/v2"
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

func parseXML(path string) (*Data, error) {
  d, err := os.ReadFile(path)
  if err != nil {
    return nil, err
  }
  
  doc, err := xmlquery.Parse(strings.NewReader(string(d)))
  doc2, err := xmlquery.Parse(strings.NewReader(string(getText(doc, "//cbc:Description"))))
  if err != nil {
    return nil, err
  }
  
  datos := &Data{
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

func procesarZip(path string) error {
  r, err := zip.OpenReader(path)
  if err != nil {
      return err
  }
  defer r.Close()

  var xmlPath, pdfPath string

  // Extraer archivos
  for _, f := range r.File {
    rc, err := f.Open()
    if err != nil {
        return err
    }
    defer rc.Close()

    outPath := filepath.Join(".", f.Name)
    outFile, err := os.Create(outPath)
    if err != nil {
        return err
    }

    _, err = io.Copy(outFile, rc)
    outFile.Close()
    if err != nil {
        return err
    }

    if strings.HasSuffix(strings.ToLower(f.Name), ".xml") {
        xmlPath = outPath
    } else if strings.HasSuffix(strings.ToLower(f.Name), ".pdf") {
        pdfPath = outPath
    }
  }

  if xmlPath != "" && pdfPath != "" {
      proveedor, err := parseXML(xmlPath)
      if err != nil {
          return err
      }

      destinoDir := "pdfs_" + proveedor.Proveedor
      os.MkdirAll(destinoDir, os.ModePerm)

      nuevoNombre := strings.ReplaceAll(proveedor.Proveedor, " ", "_") + ".pdf"
      destino := filepath.Join(destinoDir, nuevoNombre)

      if err := os.Rename(pdfPath, destino); err != nil {
          return err
      }

      fmt.Println("PDF movido a:", destino)
      fmt.Println(proveedor)
  }

  return nil
}

func esZipValido(path string) bool {
    f, err := os.Open(path)
    if err != nil {
        return false
    }
    defer f.Close()

    cabecera := make([]byte, 4)
    if _, err := f.Read(cabecera); err != nil {
        return false
    }

    return cabecera[0] == 0x50 && cabecera[1] == 0x4B
}

func main() {
  
  err := procesarZip("docs/ad08110421420002500028394.zip")
  if err != nil {
    fmt.Println(err)
  }
}