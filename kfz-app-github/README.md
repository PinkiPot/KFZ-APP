# KFZ-APP (V26.07.26)

KFZ-Verwaltungs- und Wartungsanwendung für Fahrzeuge, Teilekatalog, Wartungsintervalle und Werkzeuge.

## 📁 Projektstruktur für GitHub

```text
├── index.html          # Startseite & HTML-Struktur der Anwendung
├── README.md           # Projektdokumentation und Bereitstellungsanleitung
├── .gitignore          # Git-Ausschlussregeln
├── css/
│   └── style.css       # Komplettes Design & responsives Layout (inkl. Dark-Mode)
└── js/
    ├── data.js         # KFZ-Datenbank (aus Excel importierte Werkstatt- & Teiledaten)
    ├── app.js          # Hauptlogik der KFZ-App (Tabs, Kataloge, Benutzer, Rendering)
    └── patch-v26.js    # Erweiterungen V26 (Hotspot-Logik, Live-Speicherung, Farbverläufe)
```

## 🚀 Bereitstellung auf GitHub & GitHub Pages

1. **Repository erstellen**:
   - Erstelle auf GitHub ein neues Repository (z. B. `kfz-app`).
2. **Dateien hochladen / committen**:
   ```bash
   git init
   git add .
   git commit -m "Initial commit: KFZ-APP modularisiert"
   git branch -M main
   git remote add origin https://github.com/<DEIN-NUTZERNAME>/<REPO-NAME>.git
   git push -u origin main
   ```
3. **GitHub Pages aktivieren (Kostenloses Hosting)**:
   - Öffne dein GitHub-Repository im Browser.
   - Gehe zu **Settings** -> **Pages**.
   - Wähle unter **Branch** den Branch `main` und `/ (root)` aus und klicke auf **Save**.
   - Deine App ist nach ca. 1 Minute unter `https://<DEIN-NUTZERNAME>.github.io/<REPO-NAME>/` weltweit erreichbar.

## 🛠 Features

- **Wartungs- & Verbaulisten**: Übersichtliche Nachverfolgung von Verschleißteilen, Intervallen und Wechseldaten.
- **Interaktiver Teile- & Bilderkatalog**: Explosionszeichnungen mit klickbaren Hotspots und Teilezuordnungen.
- **Dynamische Berechnungen**: Automatische Berechnung von Resttagen, Restlaufleistung und optische Farbverläufe/Ampeln.
- **Benutzerverwaltung & Berechtigungen**: Admin-, Werkstatt- und Gastrollen.
- **Lokale & sichere Speicherung**: Automatisches Speichern im Browser (`localStorage`) ohne Datenverlust.
