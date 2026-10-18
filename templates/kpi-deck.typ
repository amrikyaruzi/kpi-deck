// KPI slide deck: one slide per indicator, data supplied by R as JSON.
// Compile (R does this for you):
//   typst compile --root <project> --input data=/build/deck.json templates/kpi-deck.typ out.pdf

#let deck = json(sys.inputs.at("data", default: "/build/deck.json"))
#let meta = deck.meta

// ---- Look and feel -----------------------------------------------------------
#let brand = rgb("#1F4E79")
#let head-grey = rgb("#A6A6A6")
#let below-red = rgb("#C00000")
#let muted = rgb("#7F7F7F")
#let line-w = 0.75pt

#set page(width: 13.333in, height: 7.5in, margin: 0pt)
#set text(font: ("Arial", "Liberation Sans", "Helvetica", "Inter", "DejaVu Sans"), size: 11pt, fill: black)
#set par(leading: 0.5em)

// Footer used on every indicator slide: logo + organisation, brand bar beneath
#let footer() = {
  place(bottom + left, rect(width: 100%, height: 0.13in, fill: brand, stroke: none))
  place(bottom + left, dx: 0.28in, dy: -0.17in,
    grid(columns: 2, column-gutter: 0.1in, align: horizon,
      image(meta.logo, height: 0.26in),
      text(weight: "bold", size: 11pt, meta.footer)))
}

#let label(t) = text(weight: "bold", t)
#let ul-label(t) = underline(evade: false, offset: 2pt, text(weight: "bold", t))

// ---- Title slide (branded) ------------------------------------------------
#let dark-red = rgb("#1F4E79")
#if meta.title_slide {
  page(fill: white)[
    #set text(font: ("Helvetica", "Arial", "Liberation Sans", "Inter", "DejaVu Sans"))
    #place(top + center, dy: 0.335in, image(meta.logo, height: 1.974in))
    #place(top + left, dx: 1.595in, dy: 2.579in,
      rect(width: 10.143in, height: 1.348in, stroke: 2.25pt + dark-red, fill: white,
        align(center + horizon, {
          let parts = meta.facility.split(",").map(p => p.trim())
          set par(leading: 14pt)
          text(size: 36pt, weight: "bold", upper(parts.at(0)))
          if parts.len() > 1 { linebreak(); text(size: 28pt, parts.slice(1).join(", ")) }
        })))
    #place(top + left, dx: 1.667in, dy: 4.09in, block(width: 10in, align(center, text(size: 18pt, meta.deck_title))))
    #place(top + left, dx: 1.667in, dy: 4.41in, block(width: 10in, align(center, text(size: 16pt, meta.subtitle))))
    #place(top + left, dx: 1.667in, dy: 4.71in, block(width: 10in, align(center, text(size: 18pt, meta.period))))
  ]
}

// ---- Slide 2: indicators in this deck ----------------------------------------------
#if meta.at("indicator_list", default: true) {
  page(fill: white)[
    #set text(font: ("Helvetica", "Arial", "Liberation Sans", "Inter", "DejaVu Sans"))
    #place(top + left, dx: 0.916in, dy: 0.228in,
      text(size: 24pt, weight: "bold", fill: rgb("#7F7F7F"))[Indicators:])
    #place(top + left, dx: 0.816in, dy: 0.733in,
      rect(width: 12in, height: 6.54in, stroke: 2.25pt + dark-red, inset: (left: 0.65in, top: 0.15in), {
        set text(size: 14pt, weight: "bold")
        set list(marker: [•], body-indent: 0.25in, spacing: 0.32in)
        list(..deck.slides.map(s => [#s.title]))
      }))
  ]
}

// ---- One slide per indicator ---------------------------------------------------
#let header-box(s) = table(
  columns: (30%, 70%),
  stroke: line-w,
  inset: (x: 7pt, y: 7pt),
  // Left: measure / target / schedule
  block(height: 1.02in)[
    #label[Measure:] #s.measure
    #v(1fr)
    #label[Target:] #s.target
    #v(1fr)
    #label[Reporting schedule:] #s.schedule
  ],
  // Right: definition + numerator/denominator as a fraction
  block(height: 1.02in)[
    #label[Definition:] #s.definition
    #v(0.16in)
    #grid(
      columns: (1.1in, auto),
      column-gutter: 6pt,
      inset: (x: 0pt, y: 2.5pt),
      align: (left + horizon, center + horizon),
      // fraction line = bottom border of the numerator cell; it spans the wider of
      // the two texts, so it also runs under "x 100%"
      stroke: (x, y) => if x == 1 and y == 0 { (bottom: 0.6pt) } else { none },
      label[Numerator:], [#s.numerator_text × 100%],
      label[Denominator:], s.denominator_text,
    )
  ],
)

#let rate-cell(cell) = if cell.below { text(fill: below-red, weight: "bold", cell.rate) } else { cell.rate }

#let data-table(s) = table(
  columns: (1.55fr,) + (1fr,) * 12,
  stroke: line-w,
  inset: (x: 5pt, y: 5.5pt),
  align: (x, y) => if x == 0 { left + horizon } else { center + horizon },
  fill: (x, y) => if y == 0 { head-grey } else { none },
  [Period], ..s.months.map(m => m.label),
  [Numerator], ..s.months.map(m => m.num),
  [Denominator], ..s.months.map(m => m.den),
  [Rate, %], ..s.months.map(rate-cell),
  [Quarterly Rate %],
  ..s.quarters.map(q => table.cell(colspan: 3, text(weight: "bold", rate-cell(q)))),
)

#let numbered(items) = if items.len() == 0 { [] } else {
  set enum(numbering: "1.", indent: 0pt, body-indent: 0.5em, spacing: 0.45em)
  enum(..items.map(i => [#i]))
}

#let notes-panel(s) = block(height: 100%, width: 100%)[
  #set text(size: 10.5pt)
  #ul-label[Analysis:]
  #if s.analysis_auto { h(4pt); text(size: 8pt, fill: muted, style: "italic")[(auto-drafted, review before circulating)] }
  #v(-2pt)
  #numbered(s.analysis)
  #v(6pt)
  #ul-label[Action:]
  #v(-2pt)
  #if s.action.len() == 0 { text(fill: muted, style: "italic")[To be completed by the responsible unit.] } else { numbered(s.action) }
  #v(1fr)
  #label[#underline(evade: false, offset: 2pt)[Data Provider]:] #s.data_provider
  #v(6pt)
  #label[#underline(evade: false, offset: 2pt)[Responsible]:] #s.responsible
]

#for s in deck.slides {
  page[
    #footer()
    #block(inset: (left: 0.24in, right: 0.24in, top: 0.2in, bottom: 0.62in), height: 100%, width: 100%)[
      #grid(
        rows: (auto, auto, auto, 1fr),
        text(size: 26pt, weight: "bold", s.title),
        v(0.12in),
        stack(header-box(s), v(-line-w), data-table(s)),
        block(height: 100%, inset: (top: -line-w), table(
          columns: (37%, 63%),
          rows: (100%,),
          stroke: line-w,
          inset: 7pt,
          notes-panel(s),
          align(center + horizon, image(s.chart, width: 100%, height: 100%, fit: "contain")),
        )),
      )
    ]
  ]
}
