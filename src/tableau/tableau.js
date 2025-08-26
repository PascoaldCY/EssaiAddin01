/* global document, Office, Word */

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
      // Bouton de démarrage
      document.getElementById("start-btn").onclick = startExercise;

      // Boutons de l'exercice
      document.getElementById("init-btn").onclick = initializeDocument;
      document.getElementById("clear-btn").onclick = clearDocument;
      document.getElementById("apply-style-btn").onclick = applyTableStyle;
      document.getElementById("validate-btn").onclick = validateTable;

      setStatus("Cliquez sur 'Commencer l'exercice' pour démarrer.");
  }
});

function setStatus(text) {
  const s = document.getElementById("status");
  if (s) s.textContent = text;
}

/** --------- CONFIRM maison --------- */
function askConfirm(message) {
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
  if (Word && Word.Alignment) {
    if (value === "center") return Word.Alignment.centered;
    if (value === "right")  return Word.Alignment.right;
    return Word.Alignment.left;
  }
  if (value === "center") return "Centered";
  if (value === "right")  return "Right";
  return "Left";
}

/** --------- Bouton de départ --------- */
async function startExercise() {
  await Word.run(async (context) => {
    const body = context.document.body;
    body.paragraphs.load("items");
    await context.sync();

    const notEmpty =
      body.paragraphs.items.length > 1 ||
      (body.paragraphs.items.length === 1 && body.paragraphs.items[0].text.trim() !== "");

    let ok = true;
    if (notEmpty) {
        ok = await askConfirm("Le document contient déjà du texte. Voulez-vous l'effacer pour commencer l'exercice ?");
        if (!ok) {
            setStatus("Exercice annulé.");
            return;
        }
        body.clear();
        await context.sync();
    }

    // Affiche le reste des contrôles et masque le bouton de départ
    document.getElementById("exercise-container").style.display = "block";
    document.getElementById("start-btn").style.display = "none";

    setStatus("Prêt ! Insérez le tableau ou appliquez un style.");
  });
}

/** --------- Actions existantes --------- */
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
    const alignUI     = document.getElementById("cell-align").value;
    const banded      = document.getElementById("banded").checked;

    await Word.run(async (context) => {
      const tables = context.document.body.tables;
      tables.load("items");
      await context.sync();

      if (tables.items.length === 0) { setStatus("Aucun tableau trouvé."); return; }
      const table = tables.items[0];

      table.rows.load("items");
      await context.sync();
      for (const row of table.rows.items) row.cells.load("items");
      await context.sync();

      const headerRow = table.rows.items[0];
      headerRow.shadingColor = headerColor;
      headerRow.font.color = "#ffffff";
      headerRow.font.bold = true;

      try {
        table.getBorder("InsideHorizontal").color = borderColor;
        table.getBorder("InsideVertical").color = borderColor;
        table.getBorder("Outside").color = borderColor;
      } catch (_) {}

      for (let i = 1; i < table.rows.items.length; i++) {
        const row = table.rows.items[i];
        if (banded && i % 2 === 1) row.shadingColor = bandColor;
        else row.shadingColor = "#ffffff";
      }

      const wordAlign = getWordAlignment(alignUI);

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

      if (tables.items.length === 0) { setStatus("❌ Aucun tableau trouvé."); return; }
      const table = tables.items[0];

      table.rows.load("items");
      await context.sync();
      for (const row of table.rows.items) row.cells.load("items");
      await context.sync();

      const rowCount = table.rows.items.length;
      const colCount = table.rows.items[0].cells.items.length;

      if (rowCount < 1 || colCount < 1) { setStatus("❌ Le tableau est vide ou mal structuré."); return; }

      const headerRow = table.rows.items[0];
      headerRow.load("shadingColor, font/color, font/bold");
      await context.sync();

      const bg = (headerRow.shadingColor || "#ffffff").toLowerCase();
      const fg = (headerRow.font.color || "#000000").toLowerCase();
      const headerOK = headerRow.font.bold && fg !== bg && !(bg === "#ffffff" && fg === "#ffffff");

      let bandsOK = true;
      if (rowCount > 1) {
        const rowsAfterHeader = table.rows.items.slice(1);
        const anyColored = rowsAfterHeader.some(
          row => row.shadingColor && row.shadingColor.toLowerCase() !== "#ffffff" && row.shadingColor.toLowerCase() !== "white"
        );
        if (anyColored) {
          for (let i = 0; i < rowsAfterHeader.length - 1; i++) {
            if (rowsAfterHeader[i].shadingColor === rowsAfterHeader[i + 1].shadingColor) {
              bandsOK = false;
              break;
            }
          }
        }
      }

      let alignmentOK = true;
      for (const row of table.rows.items) {
        for (const cell of row.cells.items) cell.body.paragraphs.load("items/alignment");
      }
      await context.sync();

      for (const row of table.rows.items) {
        for (const cell of row.cells.items) {
          for (const p of cell.body.paragraphs.items) {
            if (p.alignment === "Unknown") alignmentOK = false;
          }
        }
      }

      if (headerOK && bandsOK && alignmentOK) setStatus("✅ Tableau valide et bien mis en forme.");
      else {
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
