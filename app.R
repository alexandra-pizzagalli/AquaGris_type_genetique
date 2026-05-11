library(shiny)
library(bslib)
library(openxlsx)

# ─────────────────────────────────────────────────────────────────────────────
# Données de référence
# ─────────────────────────────────────────────────────────────────────────────

groupes_taxo <- c("",
                  "Algues et plantes aquatiques",
                  "Crustacés",
                  "Poissons d'eau douce et euryhalins",
                  "Poissons marins",
                  "Mollusques",
                  "Animaux aquatiques divers"
)

especes_par_groupe <- list(
  "Poissons marins" = c(
    "Sparus aurata – Daurade royale",
    "Dicentrarchus labrax – Bar européen",
    "Scophthalmus maximus – Turbot",
    "Solea solea – Sole commune",
    "Argyrosomus regius – Maigre",
    "Thunnus thynnus – Thon rouge",
    "Gadus morhua – Cabillaud"
  ),
  "Poissons d'eau douce et euryhalins" = c(
    "Salmo salar – Saumon atlantique",
    "Oncorhynchus mykiss – Truite arc-en-ciel",
    "Cyprinus carpio – Carpe commune",
    "Anguilla anguilla – Anguille européenne",
    "Silurus glanis – Silure glane"
  ),
  "Mollusques" = c(
    "Crassostrea gigas – Huître creuse",
    "Ostrea edulis – Huître plate",
    "Mytilus edulis – Moule commune",
    "Mytilus galloprovincialis – Moule méditerranéenne",
    "Pecten maximus – Coquille Saint-Jacques",
    "Haliotis tuberculata – Ormeau"
  ),
  "Crustacés" = c(
    "Penaeus vannamei – Crevette vannamei",
    "Penaeus monodon – Crevette géante",
    "Homarus gammarus – Homard européen",
    "Carcinus maenas – Crabe vert"
  ),
  "Algues et plantes aquatiques" = c(
    "Undaria pinnatifida – Wakame",
    "Saccharina latissima – Kombu royal",
    "Palmaria palmata – Dulse",
    "Porphyra umbilicalis – Nori",
    "Gracilaria spp.",
    "Ulva lactuca – Laitue de mer"
  ),
  "Animaux aquatiques divers" = c(
    "Apostichopus japonicus – Concombre de mer",
    "Paracentrotus lividus – Oursin violet",
    "Sepia officinalis – Seiche commune",
    "Octopus vulgaris – Pieuvre commune"
  )
)

types_elevage    <- c("", "Propagation en captivité", "Souche", "Variété", "Sauvage")
origines         <- c("", "Développé dans le pays", "Introduite", "Inconnu")
compositions_spp <- c("", "Introgressée", "Espèce pure")
methodes_selection <- c(
  "Sélection combinée", "Sélection familiale", "Sélection génomique",
  "Sélection assistée par marqueurs", "Sélection massale",
  "Sélection intra-famille", "Non connue"
)
programmes_types  <- c("", "Propriété privée", "Partenariat public-privé",
                        "Propriété publique", "Inconnu", "Autre")
strategies_commerc <- c("", "Oui – ventes commerciales", "Oui – financement public dédié",
                         "Oui – financement mixte", "Non", "En cours de développement")
mesures_conservation <- c(
  "Conservation in situ", "Conservation ex situ in vivo",
  "Conservation ex situ in vitro", "Cryoconservation", "Aucune"
)

# ─────────────────────────────────────────────────────────────────────────────
# Helpers UI
# ─────────────────────────────────────────────────────────────────────────────

lbl <- function(txt, tip = NULL, required = FALSE) {
  req_star <- if (required) tags$span(class = "required", " *") else NULL
  help     <- if (!is.null(tip))
    tags$span(class = "help-icon", title = tip, "ⓘ")
  else NULL
  tagList(txt, req_star, help)
}

section_card <- function(..., title = NULL, icon = NULL) {
  title_content <- if (!is.null(title))
    div(class = "section-title",
        if (!is.null(icon)) tagList(icon, " ") else NULL,
        title)
  else NULL
  div(class = "card-section", title_content, ...)
}

# ─────────────────────────────────────────────────────────────────────────────
# UI
# ─────────────────────────────────────────────────────────────────────────────
ui <- page_fluid(
  theme = bs_theme(
    version    = 5,
    bootswatch = "flatly",
    primary    = "#1a6b8a",
    "navbar-bg" = "#1a6b8a"
  ),

  tags$head(tags$style(HTML("
    body { background:#f0f4f7; font-family:'Segoe UI',Arial,sans-serif; font-size:0.93rem; }
    .fao-header {
      background: linear-gradient(135deg,#1a6b8a 0%,#0d4a63 100%);
      color:white; padding:16px 28px; margin-bottom:22px;
      border-radius:0 0 14px 14px;
      display:flex; align-items:center; gap:16px;
      box-shadow:0 4px 18px rgba(0,0,0,0.15);
    }
    .fao-logo-wrap {
      width:50px;height:50px;background:white;border-radius:50%;
      display:flex;align-items:center;justify-content:center;flex-shrink:0;
    }
    .fao-title   { font-size:1.35rem; font-weight:700; line-height:1.2; }
    .fao-subtitle{ font-size:0.8rem; opacity:0.82; margin-top:2px; }
    .card-section {
      background:white; border-radius:10px; padding:20px 26px;
      margin-bottom:18px; box-shadow:0 2px 8px rgba(0,0,0,0.06);
      border-left: 4px solid transparent;
      transition: border-left-color .2s;
    }
    .card-section:focus-within { border-left-color:#1a6b8a; }
    .section-title {
      color:#1a6b8a; font-weight:700; font-size:1rem;
      border-bottom:2px solid #e0ecf3; padding-bottom:8px;
      margin-bottom:14px; display:flex; align-items:center; gap:6px;
    }
    .subsection {
      background:#f5fafc; border-left:3px solid #1a6b8a;
      padding:14px 16px; border-radius:0 8px 8px 0; margin-bottom:14px;
    }
    .subsection-title { font-weight:600; color:#1a6b8a; margin-bottom:10px; font-size:0.9rem; }
    .form-label { font-weight:500; color:#333; margin-bottom:3px; }
    .help-icon  { color:#1a6b8a; font-size:0.8rem; margin-left:3px; cursor:help; vertical-align:super; }
    .required   { color:#dc3545; margin-left:2px; }
    hr.subtle   { border-color:#e8ecef; margin:14px 0; }
    .btn-export {
      background:#1a6b8a; color:white !important; border:none;
      padding:9px 24px; border-radius:8px; font-weight:600;
      font-size:0.9rem; cursor:pointer; transition:background .2s;
    }
    .btn-export:hover { background:#0d4a63; }
    .progress-wrap { margin-bottom:20px; }
    .progress-label { font-size:0.82rem; color:#555; margin-bottom:4px; font-weight:500; }
    .progress { height:8px !important; border-radius:4px; }
    .progress-bar { background:#1a6b8a !important; }
    .nav-tabs .nav-link { color:#555; font-weight:500; }
    .nav-tabs .nav-link.active { color:#1a6b8a; font-weight:700; border-bottom:3px solid #1a6b8a; }
    @media(max-width:768px){
      .card-section { padding:14px 14px; }
      .fao-header   { padding:12px 14px; }
    }
  "))),

  # ── En-tête FAO ─────────────────────────────────────────────────────────────
  div(class = "fao-header",
      div(
        div(class = "fao-title", "AquaGRIS – Formulaire Type génétique"),
        div(class = "fao-subtitle",
            "Food and Agriculture Organization of the United Nations | ",
            tags$span(style = "opacity:.7;", "Aquatic Genetic Resources Information System")
        )
      )
  ),

  div(class = "container-fluid", style = "max-width:1100px;",

      # ── Barre de progression ─────────────────────────────────────────────────
      div(class = "progress-wrap",
          div(class = "progress-label",
              textOutput("progress_label", inline = TRUE)),
          div(class = "progress",
              div(class = "progress-bar", role = "progressbar",
                  style = "width:0%", id = "prog_bar"))
      ),

      navset_tab(
        id = "main_tabs",

        # ═══════════════════════════════════════════════════════════════════════
        # ONGLET 1 – Type génétique
        # ═══════════════════════════════════════════════════════════════════════
        nav_panel("🧬 Type génétique",

                  br(),

                  # Identification du déclarant
                  section_card(title = "Identification du déclarant", icon = "🌍",
                               fluidRow(
                                 column(6,
                                        textInput("contact", lbl("NOM Prénom"), placeholder = "")
                                 ),
                                 column(6,
                                        textInput("email", lbl("Email"), placeholder = "")
                                 )
                               )
                  ),

                  # Espèce de référence
                  section_card(title = "Espèce de référence", icon = "🔬",
                               fluidRow(
                                 column(6,
                                        selectInput("groupe_taxo",
                                                    lbl("Groupe taxonomique"),
                                                    choices = groupes_taxo, selected = "")
                                 ),
                                 column(6,
                                        textInput("nom_espece",
                                                    lbl("Nom de l'espèce"),
                                                    placeholder = "")
                                 )
                               )
                  ),

                  # Identification de la lignée
                  section_card(title = "Identification de la lignée", icon = "🏷️",
                               fluidRow(
                                 column(6,
                                        textInput("nom_type",
                                                  lbl("Nom de la lignée"),
                                                  placeholder = "")
                                 ),
                                 column(6,
                                        textAreaInput("desc_type", lbl("Description"),
                                             height = "80px",
                                             placeholder = "Décrivez brièvement la lignée, ses caractéristiques…")
                                        )
                               ),
                               fluidRow(
                                 column(4,
                                        selectInput("type_primaire",
                                                    lbl(
                                                        "Classification principale du type génétique"),
                                                    choices = c("(aucun)" = "", types_elevage),
                                                    selected = "")
                                 ),
                                 column(4,
                                        numericInput("pct_production",
                                                     lbl(
                                                         "Pourcentage de la production nationale totale"),
                                                     value = NA, min = 0, max = 100, step = 1)
                                 ),
                                 column(4,
                                        selectInput("origine_type",
                                                    lbl("Origine de ce Type génétique"),
                                                    choices = c("(aucun)" = "", origines),
                                                    selected = "")
                                 )
                               ),
                               fluidRow(
                                 column(6,
                                        selectInput("risques_different",
                                                    lbl("Les risques (génétiques, sanitaires, environnementaux) associés à cette lignée diffèrent-ils de ceux de l'espèce ?"),
                                                    choices = c("(aucun)" = "",
                                                                "Risques plus importants" = "yes",
                                                                "Risques moins importants" = "no",
                                                                "Mêmes risques" = "unknown"),
                                                    selected = "")
                                 ),
                                 column(6,
                                        selectInput("composition_spp",
                                                    lbl("Composition en espèces"),
                                                    choices = c("(aucun)" = "", compositions_spp),
                                                    selected = "")
                                 )
                               )
                  ),

                  # Amélioration génétique
                  section_card(title = "Amélioration génétique", icon = "🧪",
                               checkboxGroupInput("methodes_amelio",
                                                  lbl("Ce type a-t-il été soumis à une amélioration génétique délibérée ?"),
                                                  choices = methodes_selection,
                                                  inline = TRUE),
                               hr(class = "subtle"),
                               fluidRow(
                                 column(6,
                                        selectInput("selection_ongoing",
                                                    lbl("La sélection est-elle en cours ?"),
                                                    choices = c("(aucun)" = "", "Oui" = "yes", "Non" = "no"),
                                                    selected = "")
                                 ),
                                 column(6,
                                        selectInput("type_programme",
                                                    lbl("Type de programme de sélection"),
                                                    choices = c("(aucun)" = "", programmes_types),
                                                    selected = "")
                                 ),
                                 column(6,
                                        selectInput("strategie_commerc",
                                                    lbl("Stratégie de commercialisation / financement à long terme ?"),
                                                    choices = c("(aucun)" = "", strategies_commerc),
                                                    selected = "")
                                 ),
                                 column (6,
                                         textAreaInput("details_programme", lbl("Détails sur le programme"),
                                             height = "70px",
                                             placeholder = "Objectifs…")
                                 )
                               ),
                  ),

                  # Historique et paramètres génétiques
                  section_card(title = "Historique et paramètres génétiques", icon = "📊",
                               fluidRow(
                                 column(6,
                                        selectInput("historique_doc",
                                                    lbl("Historique d'élevage documenté ? (Archives, publications, rapports internes)"),
                                                    choices = c("(aucun)" = "", "Oui" = "yes", "Non" = "no",
                                                                "Inconnu" = "unknow"),
                                                    selected = "")
                                 ),
                                 column(6,
                                        textAreaInput("ref_historique",
                                                      lbl("Références (historique d'élevage)"),
                                                      height = "70px",
                                                      placeholder = "Auteur, année, titre, lien…")
                                 )
                               ),
                               fluidRow(
                                 column(4,
                                        selectInput("ne_estimable",
                                                    lbl("Taille effective de population (Ne) estimable ?",
                                                        "Via données généalogiques ou génomiques"),
                                                    choices = c("(aucun)" = "", "Oui" = "yes", "Non" = "no"),
                                                    selected = "")
                                 ),
                                 column(4,
                                        conditionalPanel("input.ne_estimable == 'yes'",
                                                         numericInput("ne_valeur",
                                                                      lbl("Taille effective de population (Ne)"),
                                                                      value = NA, min = 1, step = 1)
                                        )
                                 ),
                                 column(4,
                                        numericInput("nb_generations",
                                                     lbl("Nombre de générations maintenues en captivité"),
                                                     value = NA, min = 0, step = 1)
                                 )
                               )
                  ),

                  # Registre et conservation
                  section_card(title = "Registre national et conservation", icon = "📋",
                               fluidRow(
                                 column(6,
                                        textAreaInput("raison_commerciale",
                                                      lbl("Principale raison de production commerciale de cette lignée"),
                                                      height = "80px",
                                                      placeholder = "ex. Croissance rapide, résistance aux maladies, faible FCR…")
                                 ),
                                 column(6,
                                        selectInput("registre_national",
                                                    lbl("Reconnu dans un registre national des types d'élevage ?"),
                                                    choices = c("(aucun)" = "", "Oui" = "yes", "Non" = "no"),
                                                    selected = ""),
                                        textInput("registre_details",
                                                  lbl("Détails registre"),
                                                  placeholder = "Numéro d'enregistrement, organisme…")
                                 )
                               ),
                               textAreaInput("sources_registre",
                                             lbl("Sources", "Références documentaires"),
                                             height = "60px",
                                             placeholder = "URLs, publications, documents officiels…"),
                               hr(class = "subtle"),
                               checkboxGroupInput("mesures_conservation",
                                                  lbl("Ce type est-il soumis à des mesures de conservation ?"),
                                                  choices = mesures_conservation,
                                                  inline = TRUE),
                               textAreaInput("details_conservation",
                                             lbl("Détails sur les mesures de conservation",
                                                 "Localisation, organisme responsable…"),
                                             height = "70px",
                                             placeholder = "Décrivez les dispositifs en place…"),
                               hr(class = "subtle"),
                               fluidRow(
                                 column(4,
                                        selectInput("export_type",
                                                    lbl("Ce type a-t-il été exporté hors de votre pays ?"),
                                                    choices = c("(aucun)" = "", "Oui" = "yes", "Non" = "no"),
                                                    selected = "")
                                 ),
                                 column(8,
                                        conditionalPanel("input.export_type == 'yes'",
                                                         textAreaInput("details_export",
                                                                       lbl("Détails sur les exportations",
                                                                           "Pays destinataires, volumes, dates"),
                                                                       height = "70px",
                                                                       placeholder = "ex. Espagne (2018–2022), Portugal (2020)…")
                                        )
                                 )
                               )
                  )
        ),

        # ═══════════════════════════════════════════════════════════════════════
        # ONGLET 2 – Récapitulatif
        # ═══════════════════════════════════════════════════════════════════════
        nav_panel("📄 Récapitulatif",
                  br(),
                  section_card(title = "Récapitulatif des données saisies", icon = "📝",
                               uiOutput("recap_html")
                  ),
                  hr(class = "subtle"),
                  section_card(title = "Télécharger et envoyer votre fiche", icon = "📧",
                               div(style = "display:flex; align-items:flex-start; gap:18px; margin-bottom:18px;",
                                   div(style = "background:#1a6b8a; color:white; font-weight:700; font-size:1.1rem;
                                      border-radius:50%; width:34px; height:34px; flex-shrink:0;
                                      display:flex; align-items:center; justify-content:center;", "1"),
                                   div(style = "flex:1;",
                                       div(style = "font-weight:600; color:#1a6b8a; margin-bottom:6px;",
                                           "Téléchargez votre fiche complète au format Excel"),
                                       div(style = "font-size:0.86rem; color:#666; margin-bottom:10px;",
                                           "Le fichier contient un onglet : Type génétique primaire."),
                                       downloadButton("export_excel",
                                                      "⬇ Télécharger la fiche Excel (.xlsx)",
                                                      class = "btn-export",
                                                      style = "font-size:0.95rem; padding:10px 26px;")
                                   )
                               ),
                               hr(class = "subtle"),
                               div(style = "display:flex; align-items:flex-start; gap:18px;",
                                   div(style = "background:#0d4a63; color:white; font-weight:700; font-size:1.1rem;
                                      border-radius:50%; width:34px; height:34px; flex-shrink:0;
                                      display:flex; align-items:center; justify-content:center;", "2"),
                                   div(style = "flex:1;",
                                       div(style = "font-weight:600; color:#1a6b8a; margin-bottom:6px;",
                                           "Envoyez le fichier par e-mail"),
                                       div(style = "background:#f0f7ff; border:1px solid #b8d6f0; border-radius:8px;
                                          padding:12px 16px; font-size:0.9rem; color:#2c5f8a;",
                                           "📎 Joignez le fichier Excel téléchargé à un e-mail et envoyez-le à ",
                                           tags$strong("alexandra.pizzagalli@ifremer.fr"),
                                           tags$br(), tags$br(),
                                           tags$span(style = "font-size:0.82rem; color:#666;",
                                                     "Objet suggéré : AquaGRIS – [Nom de l'espèce] – [Date]")
                                       )
                                   )
                               )
                  )
        )
      )
  ))

# ─────────────────────────────────────────────────────────────────────────────
# SERVER
# ─────────────────────────────────────────────────────────────────────────────
server <- function(input, output, session) {

  # ── Mise à jour dynamique des espèces selon le groupe taxonomique ────────
  observeEvent(input$groupe_taxo, {
    especes <- especes_par_groupe[[input$groupe_taxo]]
    if (is.null(especes)) especes <- character(0)
    updateSelectInput(session, "nom_espece",
                      choices = c("(aucun)" = "", especes,
                                  "Autre – saisir manuellement" = "other"))
  }, ignoreInit = FALSE)

  # ── Barre de progression ─────────────────────────────────────────────────
  progress_pct <- reactive({
    fields_remplis <- c(
      nchar(input$groupe_taxo) > 0,
      nchar(input$nom_espece) > 0,
      nchar(input$nom_type) > 0,
      nchar(input$type_primaire) > 0
    )
    round(100 * sum(fields_remplis) / length(fields_remplis))
  })

  output$progress_label <- renderText({
    paste0("Progression du formulaire : ", progress_pct(), "%")
  })

  observe({
    pct <- progress_pct()
    session$sendCustomMessage("update_progress", list(pct = pct))
  })

  tags$head(tags$script(HTML("
    Shiny.addCustomMessageHandler('update_progress', function(msg) {
      var bar = document.getElementById('prog_bar');
      if (bar) { bar.style.width = msg.pct + '%'; bar.setAttribute('aria-valuenow', msg.pct); }
    });
  ")))

  # ── Helpers ──────────────────────────────────────────────────────────────
  fmt_val <- function(x) {
    if (is.null(x) || length(x) == 0) return("—")
    v <- trimws(as.character(x))
    v <- v[!is.na(v) & nchar(v) > 0]
    if (length(v) == 0) return("—")
    paste(v, collapse = " | ")
  }

  # ── Récapitulatif HTML ────────────────────────────────────────────────────
  output$recap_html <- renderUI({
    s <- input

    fmt <- function(x, na_val = "—") {
      if (is.null(x) || length(x) == 0 || (length(x) == 1 && (is.na(x) || nchar(trimws(x)) == 0)))
        na_val
      else paste(x, collapse = ", ")
    }

    row_html <- function(label, value) {
      val_display <- if (is.null(value) || length(value) == 0 ||
                         (length(value) == 1 && (is.na(value) || nchar(trimws(as.character(value))) == 0)))
        tags$span(style = "color:#aaa; font-style:italic;", "—")
      else tags$span(paste(value, collapse = ", "))
      tags$tr(
        tags$td(style = "font-weight:600; color:#1a6b8a; width:45%; padding:5px 10px 5px 0; vertical-align:top;", label),
        tags$td(style = "padding:5px 0; vertical-align:top;", val_display)
      )
    }

    section_html <- function(title, icon_txt, ...) {
      div(style = "margin-bottom:18px;",
          div(style = "font-weight:700; font-size:1rem; color:#1a6b8a;
                       border-bottom:2px solid #e0ecf3; padding-bottom:6px; margin-bottom:10px;",
              icon_txt, " ", title),
          tags$table(style = "width:100%; border-collapse:collapse;", ...)
      )
    }

    tagList(
      div(style = "background:#1a6b8a; color:white; font-weight:700; font-size:1.05rem;
                   border-radius:8px 8px 0 0; padding:12px 18px; margin-bottom:0;",
          "🧬 Type génétique PRIMAIRE – AquaGRIS"),
      div(style = "border:1px solid #cde0ea; border-top:none; border-radius:0 0 8px 8px;
                   padding:18px 20px; margin-bottom:10px;",

          section_html("Identification du déclarant", "🌍",
                       row_html("NOM Prénom", s$contact),
                       row_html("Email", s$email)
          ),

          section_html("Espèce de référence", "🔬",
                       row_html("Groupe taxonomique", s$groupe_taxo),
                       row_html("Nom de l'espèce", s$nom_espece),
                       row_html("Nom commun officiel", s$nom_commun)
          ),

          section_html("Identification de la lignée", "🏷️",
                       row_html("Nom de la lignée", s$nom_type),
                       row_html("Description", s$desc_type),
                       row_html("Type génétique primaire", s$type_primaire),
                       row_html("% estimé de la production de l'espèce", s$pct_production),
                       row_html("Origine de ce Type génétique", s$origine_type),
                       row_html("Risques associés différents de ceux de l'espèce ?", s$risques_different),
                       row_html("Composition en espèces", s$composition_spp)
          ),

          section_html("Amélioration génétique", "🧪",
                       row_html("Méthodes d'amélioration génétique appliquées", fmt(s$methodes_amelio)),
                       row_html("La sélection est-elle en cours ?", s$selection_ongoing),
                       row_html("Type de programme de sélection", s$type_programme),
                       row_html("Stratégie de commercialisation / financement ?", s$strategie_commerc),
                       row_html("Détails sur le programme", s$details_programme)
          ),

          section_html("Historique et paramètres génétiques", "📊",
                       row_html("Historique d'élevage documenté ?", s$historique_doc),
                       row_html("Références (historique d'élevage)", s$ref_historique),
                       row_html("Taille effective de population (Ne) estimable ?", s$ne_estimable),
                       if (isTRUE(s$ne_estimable == "yes"))
                         row_html("Taille effective de population (Ne)", s$ne_valeur)
                       else NULL,
                       row_html("Nombre de générations maintenues en captivité", s$nb_generations)
          ),

          section_html("Registre national et conservation", "📋",
                       row_html("Principale raison de production commerciale", s$raison_commerciale),
                       row_html("Reconnu dans un registre national ?", s$registre_national),
                       row_html("Détails registre", s$registre_details),
                       row_html("Sources", s$sources_registre),
                       row_html("Mesures de conservation", fmt(s$mesures_conservation)),
                       row_html("Détails sur les mesures de conservation", s$details_conservation),
                       row_html("Ce type a-t-il été exporté hors de votre pays ?", s$export_type),
                       row_html("Détails sur les exportations", s$details_export)
          )
      )
    )
  })

  # ── Export Excel ───────────────────────────────────────────────────────────
  output$export_excel <- downloadHandler(
    filename = function() {
      espece_nom <- if (!is.null(input$nom_espece) && nchar(input$nom_espece) > 0)
        gsub("[^A-Za-z0-9]", "_", input$nom_espece) else "espece"
      paste0("AquaGRIS_TypeElevage_", espece_nom, "_", format(Sys.Date(), "%Y%m%d"), ".xlsx")
    },
    content = function(file) {
      s <- input

      wb <- createWorkbook()

      bleu       <- "#1a6b8a"
      bleu_clair <- "#e8f4f8"
      gris       <- "#f5f5f5"

      st_titre   <- createStyle(fontName = "Calibri", fontSize = 14, fontColour = "#FFFFFF",
                                fgFill = bleu, halign = "LEFT", valign = "center",
                                textDecoration = "bold", border = "Bottom", borderColour = bleu)
      st_section <- createStyle(fontName = "Calibri", fontSize = 10, fontColour = bleu,
                                fgFill = bleu_clair, halign = "LEFT", valign = "center",
                                textDecoration = "bold", border = "Bottom", borderColour = bleu)
      st_label   <- createStyle(fontName = "Calibri", fontSize = 10, fontColour = "#333333",
                                fgFill = gris, halign = "LEFT", valign = "top",
                                textDecoration = "bold", wrapText = TRUE)
      st_valeur  <- createStyle(fontName = "Calibri", fontSize = 10, fontColour = "#111111",
                                fgFill = "#FFFFFF", halign = "LEFT", valign = "top", wrapText = TRUE)
      st_vide    <- createStyle(fontName = "Calibri", fontSize = 10, fontColour = "#aaaaaa",
                                fgFill = "#FFFFFF", halign = "LEFT", valign = "top",
                                textDecoration = "italic", wrapText = TRUE)

      write_sheet <- function(wb, sheet_name, titre, lignes) {
        addWorksheet(wb, sheet_name, gridLines = FALSE)
        setColWidths(wb, sheet_name, cols = 1:2, widths = c(52, 60))
        r <- 1
        mergeCells(wb, sheet_name, cols = 1:2, rows = r)
        writeData(wb, sheet_name, titre, startRow = r, startCol = 1)
        addStyle(wb, sheet_name, st_titre, rows = r, cols = 1:2, gridExpand = TRUE)
        setRowHeights(wb, sheet_name, rows = r, heights = 30)
        r <- r + 1
        current_section <- ""
        for (lg in lignes) {
          if (!is.null(lg$section) && lg$section != current_section) {
            current_section <- lg$section
            r <- r + 1
            mergeCells(wb, sheet_name, cols = 1:2, rows = r)
            writeData(wb, sheet_name, current_section, startRow = r, startCol = 1)
            addStyle(wb, sheet_name, st_section, rows = r, cols = 1:2, gridExpand = TRUE)
            setRowHeights(wb, sheet_name, rows = r, heights = 22)
            r <- r + 1
          }
          val <- fmt_val(lg$valeur)
          writeData(wb, sheet_name, lg$label, startRow = r, startCol = 1)
          writeData(wb, sheet_name, val,       startRow = r, startCol = 2)
          addStyle(wb, sheet_name, st_label,  rows = r, cols = 1)
          if (val == "—") {
            addStyle(wb, sheet_name, st_vide, rows = r, cols = 2)
          } else {
            addStyle(wb, sheet_name, st_valeur, rows = r, cols = 2)
          }
          setRowHeights(wb, sheet_name, rows = r, heights = 18)
          r <- r + 1
        }
      }

      lignes_type <- list(
        list(section = "Identification du déclarant", label = "NOM Prénom", valeur = s$contact),
        list(section = NULL,                           label = "Email",      valeur = s$email),
        list(section = "Espèce de référence",          label = "Groupe taxonomique",   valeur = s$groupe_taxo),
        list(section = NULL,                           label = "Nom de l'espèce",      valeur = s$nom_espece),
        list(section = NULL,                           label = "Nom commun officiel",   valeur = s$nom_commun),
        list(section = "Identification de la lignée",  label = "Nom de la lignée",                                   valeur = s$nom_type),
        list(section = NULL,                           label = "Description",                                         valeur = s$desc_type),
        list(section = NULL,                           label = "Type génétique primaire",                            valeur = s$type_primaire),
        list(section = NULL,                           label = "% estimé de la production de l'espèce",              valeur = s$pct_production),
        list(section = NULL,                           label = "Origine de ce Type génétique",                       valeur = s$origine_type),
        list(section = NULL,                           label = "Risques associés différents de ceux de l'espèce ?",  valeur = s$risques_different),
        list(section = NULL,                           label = "Composition en espèces",                             valeur = s$composition_spp),
        list(section = "Amélioration génétique",       label = "Méthodes d'amélioration génétique appliquées",       valeur = fmt_val(s$methodes_amelio)),
        list(section = NULL,                           label = "La sélection est-elle en cours ?",                   valeur = s$selection_ongoing),
        list(section = NULL,                           label = "Type de programme de sélection",                     valeur = s$type_programme),
        list(section = NULL,                           label = "Stratégie de commercialisation / financement ?",     valeur = s$strategie_commerc),
        list(section = NULL,                           label = "Détails sur le programme",                           valeur = s$details_programme),
        list(section = "Historique et paramètres génétiques", label = "Historique d'élevage documenté ?",           valeur = s$historique_doc),
        list(section = NULL,                           label = "Références (historique d'élevage)",                  valeur = s$ref_historique),
        list(section = NULL,                           label = "Taille effective de population (Ne) estimable ?",    valeur = s$ne_estimable)
      )
      if (isTRUE(s$ne_estimable == "yes")) {
        lignes_type <- c(lignes_type, list(
          list(section = NULL, label = "Taille effective de population (Ne)", valeur = s$ne_valeur)
        ))
      }
      lignes_type <- c(lignes_type, list(
        list(section = NULL,                                label = "Nombre de générations maintenues en captivité",   valeur = s$nb_generations),
        list(section = "Registre national et conservation", label = "Principale raison de production commerciale",     valeur = s$raison_commerciale),
        list(section = NULL,                                label = "Reconnu dans un registre national ?",             valeur = s$registre_national),
        list(section = NULL,                                label = "Détails registre",                                valeur = s$registre_details),
        list(section = NULL,                                label = "Sources",                                         valeur = s$sources_registre),
        list(section = NULL,                                label = "Mesures de conservation",                         valeur = fmt_val(s$mesures_conservation)),
        list(section = NULL,                                label = "Détails sur les mesures de conservation",         valeur = s$details_conservation),
        list(section = NULL,                                label = "Ce type a-t-il été exporté hors de votre pays ?", valeur = s$export_type),
        list(section = NULL,                                label = "Détails sur les exportations",                    valeur = s$details_export)
      ))

      write_sheet(wb, "Type génétique", "Type génétique PRIMAIRE – AquaGRIS", lignes_type)
      saveWorkbook(wb, file, overwrite = TRUE)
    }
  )
}

`%||%` <- function(a, b) if (!is.null(a) && length(a) > 0 && !is.na(a[1]) && nchar(a[1]) > 0) a else b

shinyApp(ui, server)
