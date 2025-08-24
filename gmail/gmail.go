package gmail

import (
  "bufio"
  "fmt"
  "io"
  "log"
  "os"
  "strings"
  "sync"
  "time"

  "github.com/emersion/go-imap"
  "github.com/emersion/go-imap/client"
  "github.com/emersion/go-message"
  "github.com/emersion/go-message/charset"
  "github.com/emersion/go-message/mail"
)

func DescargarZips() {
  reader := bufio.NewReader(os.Stdin)

  fmt.Print("Correo Gmail: ")
  email, _ := reader.ReadString('\n')
  email = strings.TrimSpace(email)

  fmt.Print("Clave de aplicación: ")
  password, _ := reader.ReadString('\n')
  password = strings.TrimSpace(password)

  fmt.Print("Año a buscar: ")
  yearStr, _ := reader.ReadString('\n')
  yearStr = strings.TrimSpace(yearStr)

  year, err := time.Parse("2006", yearStr)
  if err != nil {
    log.Fatalf("Error en el año: %v", err)
  }

  if _, err := os.Stat("docs"); os.IsNotExist(err) {
    if err := os.Mkdir("docs", 0755); err != nil {
      log.Fatalf("No se pudo crear carpeta docs: %v", err)
    }
  }

  message.CharsetReader = charset.Reader

  //numWorkers := 12
  results := make(chan string, 1000)
  var wg sync.WaitGroup

  // lanzar los 12 workers
  for month := 1; month <= 12; month++ {
    wg.Add(1)
    go workerMes(email, password, year.Year(), month, results, &wg)
  }

  // gorutina que cierra el canal cuando todos terminan
  go func() {
    wg.Wait()
    close(results)
  }()

  // leer resultados hasta que se cierre el canal
  for res := range results {
    log.Println(res)
  }

  log.Println("✅ Descarga completa en carpeta docs/")
}

func workerMes(email, password string, year, month int, results chan<- string, wg *sync.WaitGroup) {
  defer wg.Done()

  c, err := client.DialTLS("imap.gmail.com:993", nil)
  if err != nil {
    results <- fmt.Sprintf("Worker mes %d no pudo conectarse: %v", month, err)
    return
  }
  defer c.Logout()

  if err := c.Login(email, password); err != nil {
    results <- fmt.Sprintf("Worker mes %d error login: %v", month, err)
    return
  }

  _, err = c.Select("INBOX", false)
  if err != nil {
    results <- fmt.Sprintf("Worker mes %d no pudo seleccionar INBOX: %v", month, err)
    return
  }

  section := &imap.BodySectionName{}
  since := time.Date(year, time.Month(month), 1, 0, 0, 0, 0, time.UTC)
  before := since.AddDate(0, 1, 0)

  criteria := imap.NewSearchCriteria()
  criteria.Since = since
  criteria.Before = before

  ids, err := c.Search(criteria)
  if err != nil {
    results <- fmt.Sprintf("Worker mes %d error en búsqueda: %v", month, err)
    return
  }
  results <- fmt.Sprintf("Worker mes %d encontró %d mensajes (%s - %s)", month, len(ids), since.Format("2006-01-02"), before.Format("2006-01-02"))

  for _, msgID := range ids {
    seqset := new(imap.SeqSet)
    seqset.AddNum(msgID)

    messages := make(chan *imap.Message, 1)
    go func() {
        _ = c.Fetch(seqset, []imap.FetchItem{imap.FetchEnvelope, imap.FetchBodyStructure, section.FetchItem()}, messages)
    }()

    for msg := range messages {
      if msg == nil {
        continue
      }

      r := msg.GetBody(section)
      if r == nil {
        continue
      }

      mr, err := mail.CreateReader(r)
      if err != nil {
        results <- fmt.Sprintf("❌ Worker mes %d error creando reader UID %d: %v", month, msgID, err)
        continue
      }

      for {
        p, err := mr.NextPart()
        if err != nil {
          break
        }
        if h, ok := p.Header.(*mail.AttachmentHeader); ok {
            filename, _ := h.Filename()
          if strings.HasSuffix(strings.ToLower(filename), ".zip") {
            path := fmt.Sprintf("docs/%s", filename)
            f, err := os.Create(path)
            if err != nil {
              results <- fmt.Sprintf("❌ Worker mes %d no pudo crear archivo %s: %v", month, filename, err)
              continue
            }
            writer := bufio.NewWriter(f)
            n, err := io.Copy(writer, p.Body)
            writer.Flush()
            f.Close()
            if err != nil && err != io.EOF {
              results <- fmt.Sprintf("❌ Worker mes %d error escribiendo %s: %v", month, filename, err)
              continue
            }
            results <- fmt.Sprintf("📥 Worker mes %d descargó ZIP: %s (%d bytes)", month, filename, n)
          }
        }
      }
    }
  }
}