# Japan 2026 Travel Planner

An offline, mobile-first reference website generated from the source itinerary workbook.

Open `index.html` in a modern browser. No network connection, external font, library, server, or API is required.

## Source and data pipeline

The authoritative source workbook remains at `source/Japan 2026 - Day itinerary (Organised).xlsx`.

`js/data.js` is generated from that workbook by `extract_data.py`. If the workbook changes, run the extractor with Python and reopen `index.html`.

The application keeps its data separate from the presentation layer:

- `js/data.js` holds the structured itinerary, transport, accommodation, luggage-forwarding, and Japanese-reference data.
- `js/app.js` renders and navigates the offline application.
- `css/style.css` provides the responsive, mobile-first interface.
