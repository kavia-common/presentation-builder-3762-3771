import React, { useId, useMemo, useState } from "react";
import PptxGenJS from "pptxgenjs";
import "./App.css";

function safeFileNamePart(value) {
  const trimmed = String(value || "").trim();
  if (!trimmed) return "presentation";
  return trimmed
    .replace(/[^\w\- ]+/g, "")
    .replace(/\s+/g, "-")
    .toLowerCase()
    .slice(0, 64);
}

// PUBLIC_INTERFACE
function App() {
  /** Single-page UI to generate a basic PPTX in the browser and download it. */
  const titleId = useId();
  const bodyId = useId();

  const [title, setTitle] = useState("Hello world");
  const [body, setBody] = useState("Hello world");
  const [status, setStatus] = useState("idle"); // idle | generating | success | error
  const [errorMessage, setErrorMessage] = useState("");
  const [lastGeneratedName, setLastGeneratedName] = useState("");

  const canGenerate = useMemo(() => {
    return status !== "generating" && (title.trim().length > 0 || body.trim().length > 0);
  }, [status, title, body]);

  // PUBLIC_INTERFACE
  const generatePpt = async () => {
    /**
     * Creates a simple one-slide PPTX with a title and body text and triggers a browser download.
     * Uses PptxGenJS client-side; no backend required.
     */
    setStatus("generating");
    setErrorMessage("");
    setLastGeneratedName("");

    try {
      const pptx = new PptxGenJS();

      // Simple theme-ish defaults (PptxGenJS has limited theming; we style the slide elements).
      pptx.layout = "LAYOUT_WIDE"; // 13.333 x 7.5 in (16:9)
      pptx.author = "PPT Generator";

      const slide = pptx.addSlide();

      // Ocean Professional palette
      const COLORS = {
        primary: "1E3A8A",
        secondary: "F59E0B",
        background: "F3F4F6",
        surface: "FFFFFF",
        text: "111827"
      };

      // Background
      slide.background = { color: COLORS.background };

      // Card surface (classic centered panel)
      slide.addShape(pptx.ShapeType.roundRect, {
        x: 0.75,
        y: 0.8,
        w: 11.83,
        h: 5.9,
        fill: { color: COLORS.surface },
        line: { color: "E5E7EB", width: 1 },
        radius: 12
      });

      // Accent bar
      slide.addShape(pptx.ShapeType.rect, {
        x: 0.75,
        y: 0.8,
        w: 11.83,
        h: 0.12,
        fill: { color: COLORS.secondary },
        line: { color: COLORS.secondary }
      });

      const titleText = title.trim() || "Untitled";
      const bodyText = body.trim() || "";

      // Title
      slide.addText(titleText, {
        x: 1.2,
        y: 1.25,
        w: 10.9,
        h: 1.0,
        fontFace: "Calibri",
        fontSize: 38,
        bold: true,
        color: COLORS.primary
      });

      // Body
      slide.addText(bodyText, {
        x: 1.2,
        y: 2.35,
        w: 10.9,
        h: 3.8,
        fontFace: "Calibri",
        fontSize: 20,
        color: COLORS.text,
        valign: "top"
      });

      // Footer note (subtle)
      slide.addText("Generated locally in your browser", {
        x: 1.2,
        y: 6.25,
        w: 10.9,
        h: 0.4,
        fontFace: "Calibri",
        fontSize: 12,
        color: "6B7280"
      });

      const name = `${safeFileNamePart(titleText)}.pptx`;
      await pptx.writeFile({ fileName: name });

      setLastGeneratedName(name);
      setStatus("success");
    } catch (err) {
      const message =
        err instanceof Error
          ? err.message
          : "Unknown error while generating the PowerPoint file.";
      setErrorMessage(message);
      setStatus("error");
    }
  };

  // PUBLIC_INTERFACE
  const onSubmit = (e) => {
    /** Form submit handler to allow Enter key generation and consistent accessibility. */
    e.preventDefault();
    if (!canGenerate) return;
    void generatePpt();
  };

  return (
    <div className="App">
      <header className="TopBar" role="banner">
        <div className="TopBar__inner">
          <div className="Brand">
            <div className="Brand__mark" aria-hidden="true" />
            <div className="Brand__text">
              <div className="Brand__title">PowerPoint Generator</div>
              <div className="Brand__subtitle">Ocean Professional • Classic</div>
            </div>
          </div>
          <div className="TopBar__hint" aria-hidden="true">
            Client-side • No backend
          </div>
        </div>
      </header>

      <main className="Page" role="main">
        <section className="Card" aria-label="PowerPoint content form">
          <div className="Card__header">
            <h1 className="Card__title">Create a simple slide</h1>
            <p className="Card__description">
              Enter a title and body text, then generate a <span className="Mono">.pptx</span> file in your
              browser.
            </p>
          </div>

          <form className="Form" onSubmit={onSubmit}>
            <div className="Field">
              <label className="Label" htmlFor={titleId}>
                Title
              </label>
              <input
                id={titleId}
                className="Input"
                type="text"
                value={title}
                onChange={(e) => setTitle(e.target.value)}
                placeholder="Hello world"
                autoComplete="off"
              />
            </div>

            <div className="Field">
              <label className="Label" htmlFor={bodyId}>
                Body text
              </label>
              <textarea
                id={bodyId}
                className="Textarea"
                value={body}
                onChange={(e) => setBody(e.target.value)}
                placeholder="Hello world"
                rows={5}
              />
            </div>

            <div className="Actions">
              <button className="Button Button--primary" type="submit" disabled={!canGenerate}>
                {status === "generating" ? "Generating…" : "Generate PPT"}
              </button>

              <button
                className="Button Button--secondary"
                type="button"
                onClick={() => {
                  setTitle("Hello world");
                  setBody("Hello world");
                  setStatus("idle");
                  setErrorMessage("");
                  setLastGeneratedName("");
                }}
              >
                Reset
              </button>
            </div>

            <div className="Status" aria-live="polite" aria-atomic="true">
              {status === "idle" && <span className="Status__muted">Ready.</span>}

              {status === "generating" && (
                <span className="Status__muted">
                  Generating your PowerPoint… this may take a moment.
                </span>
              )}

              {status === "success" && (
                <div className="Status__success" role="status">
                  <div className="Status__headline">Success</div>
                  <div className="Status__detail">
                    Download started{lastGeneratedName ? (
                      <>
                        : <span className="Mono">{lastGeneratedName}</span>
                      </>
                    ) : (
                      "."
                    )}
                  </div>
                </div>
              )}

              {status === "error" && (
                <div className="Status__error" role="alert">
                  <div className="Status__headline">Couldn’t generate PPT</div>
                  <div className="Status__detail">
                    {errorMessage || "Please try again."}
                  </div>
                </div>
              )}
            </div>
          </form>
        </section>
      </main>

      <footer className="Footer" role="contentinfo">
        <div className="Footer__inner">
          Tip: Use Tab/Shift+Tab to navigate. Press Enter on “Generate PPT” to download.
        </div>
      </footer>
    </div>
  );
}

export default App;
