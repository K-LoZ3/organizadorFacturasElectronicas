package main

import(
 "fmt"
 "os"
 "strings"
 "github.com/antchfx/xmlquery"
 //"github.com/nwaples/rardecode"
 //_ "github.com/xuri/excelize/v2"
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

func main() {
  fmt.Println(parseXML("docs/ad08000699330212360124868.xml"))
}