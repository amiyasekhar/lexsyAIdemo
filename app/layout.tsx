import "./globals.css";

export const metadata = {
  title: "LexsyAI — Document Placeholder Agent",
  description:
    "Upload a legal draft, detect placeholders, fill them conversationally, and download the completed document."
};

export default function RootLayout({
  children
}: Readonly<{
  children: React.ReactNode;
}>) {
  return (
    <html lang="en">
      <body>
        <div className="shell">
          <header className="header">
            <div className="brand">
              <div className="logo">LA</div>
              <div>
                <div className="title">LexsyAI</div>
                <div className="subtitle">Document Placeholder Agent</div>
              </div>
            </div>
            <a
              className="link"
              href="https://github.com/"
              target="_blank"
              rel="noreferrer"
            >
              Docs
            </a>
          </header>
          <main className="main">{children}</main>
          <footer className="footer">
            <span>
              Tip: Use placeholders like <code>{"{{ClientName}}"}</code> or{" "}
              <code>{"[[EffectiveDate]]"}</code> or <code>{"<<Fee>>"}</code>.
            </span>
          </footer>
        </div>
      </body>
    </html>
  );
}


