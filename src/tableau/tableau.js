/* global document, Office, Word */

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("init-btn").onclick = initializeDocument;
    document.getElementById("clear-btn").onclick = clearDocument;
    document.getElementById("apply-style-btn").onclick = applyTableStyle;
    document.getElementById("validate-btn").onclick = validateTable;
    setStatus("Add-in prêt. Cliquez sur Initialiser pour insérer le document et le tableau.");
  }
});

function setStatus(text) {
  const s = document.getElementById("status");
  if (s) s.textContent = text;
}

/** --------- CONFIRM maison (fallback si window.confirm n'est pas dispo) --------- */
function askConfirm(message) {
  try {
    if (typeof window.confirm === "function") {
      return Promise.resolve(window.confirm(message));
    }
  } catch (e) { /* ignore, fallback */ }

  return new Promise((resolve) => {
    const panel = document.getElementById("confirm-panel");
    const text = document.getElementById("confirm-text");
    const yes = document.getElementById("confirm-yes");
    const no = document.getElementById("confirm-no");

    text.textContent = message;
    panel.hidden = false;

    const cleanup = () => {
      yes.onclick = null;
      no.onclick = null;
      panel.hidden = true;
    };
    yes.onclick = () => { cleanup(); resolve(true); };
    no.onclick  = () => { cleanup(); resolve(false); };
  });
}

/** --------- Helpers --------- */
function getWordAlignment(value) {
  // map des valeurs UI -> enum Word
  if (Word && Word.Alignment) {
    if (value === "center") return Word.Alignment.centered;
    if (value === "right")  return Word.Alignment.right;
    return Word.Alignment.left;
  }
  // fallback string
  if (value === "center") return "Centered";
  if (value === "right")  return "Right";
  return "Left";
}

/** --------- Actions --------- */
async function initializeDocument() {
  try {
    await Word.run(async (context) => {
      const body = context.document.body;
      body.paragraphs.load("items");
      await context.sync();

      const notEmpty =
        body.paragraphs.items.length > 1 ||
        (body.paragraphs.items.length === 1 && body.paragraphs.items[0].text.trim() !== "");

      if (notEmpty) {
        const ok = await askConfirm("La page contient déjà du contenu. Voulez-vous l'effacer ?");
        if (!ok) { setStatus("Initialisation annulée."); return; }
        body.clear();
        await context.sync();
      }

      const data = [
        ["Nom", "Âge", "Ville"],
        ["Jean Dupont", "30", "Paris"],
        ["Marie Curie", "35", "Lyon"],
        ["Ali Benz", "28", "Marseille"],
        ["Sophie Martin", "42", "Toulouse"]
      ];

      const table = body.insertTable(data.length, data[0].length, Word.InsertLocation.end, data);
      table.styleBuiltIn = "TableGrid";

      // Charger lignes pour accéder à l'en-tête
      table.rows.load("items");
      await context.sync();

      const headerRow = table.rows.items[0];
      headerRow.font.bold = true;
      headerRow.font.color = "white";
      headerRow.shadingColor = "#0078d7";

      await context.sync();
      setStatus("Document et tableau insérés. Vous pouvez appliquer un style et valider.");
    });
  } catch (err) {
    console.error(err);
    setStatus("Erreur lors de l'initialisation : " + (err.message || err));
  }
}

async function applyTableStyle() {
  try {
    const headerColor = document.getElementById("header-color").value;
    const bandColor   = document.getElementById("band-color").value;
    const borderColor = document.getElementById("border-color").value;
    const alignUI     = document.getElementById("cell-align").value; // left | center | right
    const banded      = document.getElementById("banded").checked;

    await Word.run(async (context) => {
      const tables = context.document.body.tables;
      tables.load("items");
      await context.sync();

      if (tables.items.length === 0) {
        setStatus("Aucun tableau trouvé.");
        return;
      }

      const table = tables.items[0];

      // Charger lignes & cellules
      table.rows.load("items");
      await context.sync();
      for (const row of table.rows.items) row.cells.load("items");
      await context.sync();

      // ------ En-tête ------
      const headerRow = table.rows.items[0];
      headerRow.shadingColor = headerColor;
      headerRow.font.color = "#ffffff";
      headerRow.font.bold = true;

      // ------ Bordures (si supporté) ------
      try {
        table.getBorder("InsideHorizontal").color = borderColor;
        table.getBorder("InsideVertical").color = borderColor;
        table.getBorder("Outside").color = borderColor;
      } catch (_) { /* certaines versions ne supportent pas getBorder */ }

      // ------ Bandes alternées ------
      for (let i = 1; i < table.rows.items.length; i++) {
        const row = table.rows.items[i];
        if (banded && i % 2 === 1) {
          row.shadingColor = bandColor;        // bande
        } else {
          row.shadingColor = "#ffffff";        // "désactiver" = blanc
        }
      }

      // ------ Alignement horizontal du texte dans CHAQUE cellule ------
      const wordAlign = getWordAlignment(alignUI);

      // charger les paragraphes de toutes les cellules
      for (const row of table.rows.items) {
        for (const cell of row.cells.items) {
          cell.body.paragraphs.load("items");
        }
      }
      await context.sync();

      for (const row of table.rows.items) {
        for (const cell of row.cells.items) {
          for (const p of cell.body.paragraphs.items) {
            p.alignment = wordAlign;
          }
        }
      }

      await context.sync();
      setStatus("Style appliqué (en-tête, bordures, bandes, alignement).");
    });
  } catch (err) {
    console.error(err);
    setStatus("Erreur style : " + (err.message || err));
  }
}

async function validateTable() {
  try {
    await Word.run(async (context) => {
      const tables = context.document.body.tables;
      tables.load("items");
      await context.sync();

      if (tables.items.length === 0) {
        setStatus("❌ Aucun tableau trouvé dans le document.");
        return;
      }

      const table = tables.items[0];

      // Charger les lignes
      table.rows.load("items");
      await context.sync();

      // Charger les cellules de chaque ligne
      for (const row of table.rows.items) {
        row.cells.load("items");
      }
      await context.sync();

      const rowCount = table.rows.items.length;
      const colCount = table.rows.items[0].cells.items.length;

      if (rowCount < 1 || colCount < 1) {
        setStatus("❌ Le tableau est vide ou mal structuré.");
        return;
      }

      // Vérification en-tête
      const headerRow = table.rows.items[0];
      headerRow.load("shadingColor, font/color, font/bold");
      await context.sync();

      const bg = (headerRow.shadingColor || "#ffffff").toLowerCase();
      const fg = (headerRow.font.color || "#000000").toLowerCase();

      // L'en-tête est OK si :
      // 1. Il est en gras
      // 2. La couleur de la police est différente du fond
      // 3. Ce n'est pas blanc sur blanc
      const headerOK = headerRow.font.bold && fg !== bg && !(bg === "#ffffff" && fg === "#ffffff");

      // Vérification bandes alternées
      let bandsOK = true;
      if (rowCount > 1) { // au moins une ligne après l'en-tête
        const rowsAfterHeader = table.rows.items.slice(1);

        // Vérifie si au moins une ligne a une couleur différente de blanc
        const anyColored = rowsAfterHeader.some(
          row => row.shadingColor && row.shadingColor.toLowerCase() !== "#ffffff" && row.shadingColor.toLowerCase() !== "white"
        );

        if (anyColored) {
          // Si des couleurs sont appliquées, vérifier alternance
          for (let i = 0; i < rowsAfterHeader.length - 1; i++) {
            if (rowsAfterHeader[i].shadingColor === rowsAfterHeader[i + 1].shadingColor) {
              bandsOK = false;
              break;
            }
          }
        } else {
          // Pas de couleurs = bandes non appliquées volontairement → OK
          bandsOK = true;
        }
      }

      // Vérification alignement
      let alignmentOK = true;
      for (const row of table.rows.items) {
        for (const cell of row.cells.items) {
          cell.body.paragraphs.load("items/alignment");
        }
      }
      await context.sync();

      for (const row of table.rows.items) {
        for (const cell of row.cells.items) {
          for (const p of cell.body.paragraphs.items) {
            if (p.alignment === "Unknown") alignmentOK = false;
          }
        }
      }

      // Affichage du résultat
      if (headerOK && bandsOK && alignmentOK) {
        setStatus("✅ Tableau valide et bien mis en forme.");
      } else {
        const problems = [];
        if (!headerOK) problems.push("En-tête illisible ou non gras");
        if (!bandsOK) problems.push("Pas de bandes alternées correctes");
        if (!alignmentOK) problems.push("Alignement incohérent");
        setStatus("⚠️ Tableau trouvé mais problèmes : " + problems.join(", "));
      }
    });
  } catch (err) {
    setStatus("Erreur validation : " + (err.message || err));
  }
}

async function clearDocument() {
  try {
    const doClear = await askConfirm("Voulez-vous effacer tout le contenu du document ?");
    if (!doClear) { setStatus("Suppression annulée."); return; }

    await Word.run(async (context) => {
      context.document.body.clear();
      await context.sync();
      setStatus("Document effacé.");
    });
  } catch (err) {
    console.error(err);
    setStatus("Erreur suppression : " + (err.message || err));
  }
}