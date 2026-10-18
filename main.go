// DocStudio is a local document editor for PDF, DOCX and text files.
// It runs a small HTTP server on 127.0.0.1 and opens the UI in the browser.
package main

import (
	"embed"
	"flag"
	"fmt"
	"io/fs"
	"log"
	"net"
	"net/http"
	"os"
	"os/exec"
	"path/filepath"
	"runtime"
	"strings"
	"sync"
	"time"

	"docstudio/internal/docx"
	"docstudio/internal/fsapi"
	"docstudio/internal/pdfops"
)

//go:embed all:web
var webFS embed.FS

func main() {
	port := flag.Int("port", 0, "port to listen on (0 = pick a free port)")
	noBrowser := flag.Bool("no-browser", false, "do not open the browser")
	flag.Parse()

	ln, err := net.Listen("tcp", fmt.Sprintf("127.0.0.1:%d", *port))
	if err != nil {
		log.Fatal(err)
	}
	addr := ln.Addr().String()

	mux := http.NewServeMux()
	fsapi.Register(mux)
	docx.Register(mux)
	pdfops.Register(mux)
	mux.HandleFunc("POST /api/ping", heartbeat())

	static, _ := fs.Sub(webFS, "web")
	mux.Handle("/", http.FileServer(http.FS(static)))

	url := "http://" + addr + "/"
	// A file passed on the command line is opened on start.
	if flag.NArg() > 0 {
		if abs, err := filepath.Abs(flag.Arg(0)); err == nil {
			url += "#open=" + strings.ReplaceAll(filepath.ToSlash(abs), " ", "%20")
		}
	}
	log.Printf("DocStudio running at %s", url)
	if !*noBrowser {
		openBrowser(url)
	}
	log.Fatal(http.Serve(ln, localOnly(addr, mux)))
}

// localOnly rejects requests that do not come from this app's own origin,
// so other websites cannot read or write local files through the API.
func localOnly(addr string, next http.Handler) http.Handler {
	_, port, _ := net.SplitHostPort(addr)
	allowed := map[string]bool{"127.0.0.1:" + port: true, "localhost:" + port: true}
	return http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		if !allowed[r.Host] {
			http.Error(w, "forbidden host", http.StatusForbidden)
			return
		}
		if o := r.Header.Get("Origin"); o != "" && !allowed[strings.TrimPrefix(strings.TrimPrefix(o, "http://"), "https://")] {
			http.Error(w, "forbidden origin", http.StatusForbidden)
			return
		}
		w.Header().Set("Cache-Control", "no-store")
		next.ServeHTTP(w, r)
	})
}

// heartbeat returns a handler the UI pings every 10s. Once a window has
// pinged, the process exits after 3 minutes without pings (all windows closed).
func heartbeat() http.HandlerFunc {
	var mu sync.Mutex
	var last time.Time
	go func() {
		for range time.Tick(5 * time.Second) {
			mu.Lock()
			idle := !last.IsZero() && time.Since(last) > 3*time.Minute
			mu.Unlock()
			if idle {
				log.Print("no open windows, exiting")
				os.Exit(0)
			}
		}
	}()
	return func(w http.ResponseWriter, r *http.Request) {
		mu.Lock()
		last = time.Now()
		mu.Unlock()
		w.WriteHeader(http.StatusNoContent)
	}
}

// openBrowser opens url in an app-style window when Edge or Chrome is
// available, otherwise in the default browser.
func openBrowser(url string) {
	if runtime.GOOS == "windows" {
		for _, p := range []string{
			os.Getenv("ProgramFiles(x86)") + `\Microsoft\Edge\Application\msedge.exe`,
			os.Getenv("ProgramFiles") + `\Microsoft\Edge\Application\msedge.exe`,
			os.Getenv("ProgramFiles") + `\Google\Chrome\Application\chrome.exe`,
		} {
			if _, err := os.Stat(p); err == nil {
				if exec.Command(p, "--app="+url).Start() == nil {
					return
				}
			}
		}
		exec.Command("rundll32", "url.dll,FileProtocolHandler", url).Start()
		return
	}
	cmd := "xdg-open"
	if runtime.GOOS == "darwin" {
		cmd = "open"
	}
	exec.Command(cmd, url).Start()
}
