/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
  }
});

Office.onReady(() => {
  document.getElementById("init-btn").onclick = initializeDocument;
  document.getElementById("clear-btn").onclick = clearDocument;
  document.getElementById("apply-style-btn").onclick = applyTableStyle;
  document.getElementById("validate-btn").onclick = validateTable;
  setStatus("Add-in prêt. Cliquez sur Initialiser pour insérer le document et le tableau.");
});

function setStatus(text) {
  const s = document.getElementById("status");
  if (s) s.textContent = text;
}

async function initializeDocument() {
  try {
    await Word.run(async (context) => {
      const body = context.document.body;
      body.paragraphs.load("items");
      await context.sync();

      const notEmpty = body.paragraphs.items.length > 1 || (body.paragraphs.items.length === 1 && body.paragraphs.items[0].text.trim() !== "");
      if (notEmpty) {
        const ok = confirm("La page contient déjà du contenu. Voulez-vous l'effacer et créer un nouveau document de test ?");
        if (!ok) { setStatus("Initialisation annulée."); return; }
        body.clear();
        await context.sync();
      }

      // Crée quelques paragraphes pour simuler 1-2 pages (simple)
      body.insertParagraph("Document de test - exercice tableau", Word.InsertLocation.start).font.size = 18;
      body.insertParagraph("", Word.InsertLocation.end);
      // Données du tableau
      const data = [
        ["Nom", "Âge", "Ville"],
        ["Jean Dupont", "30", "Paris"],
        ["Marie Curie", "35", "Lyon"],
        ["Ali Benz", "28", "Marseille"],
        ["Sophie Martin", "42", "Toulouse"]
      ];
      const table = context.document.body.insertTable(data.length, data[0].length, Word.InsertLocation.end, data);
      // style de base
      table.styleBuiltIn = "TableGrid";
      table.getRow(0).font.bold = true;
      table.getRow(0).font.color = "white";
      table.getRow(0).shadingColor = "#0078d7";
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
    const bandColor = document.getElementById("band-color").value;
    const borderColor = document.getElementById("border-color").value;
    const align = document.getElementById("cell-align").value;
    const banded = document.getElementById("banded").checked;

    await Word.run(async (context) => {
      const tables = context.document.body.tables;
      tables.load("items");
      await context.sync();

      if (tables.items.length === 0) {
        setStatus("Aucun tableau trouvé dans le document.");
        return;
      }
      const table = tables.items[0];

      // en-tête
      table.getRow(0).shadingColor = headerColor;
      table.getRow(0).font.color = "#ffffff";
      table.getRow(0).font.bold = true;

      // bordures (appliquer couleur bordure simple)
      try {
        table.getBorder("InsideHorizontal").color = borderColor;
        table.getBorder("InsideVertical").color = borderColor;
        table.getBorder("Outside").color = borderColor;
      } catch (e) {
        // certaines versions peuvent ne pas supporter getBorder - ignore
      }

      // alignement et bandes
      table.rows.load("items");
      await context.sync();
      table.rows.items.forEach((r, idx) => {
        r.alignment = align; // left/center/right
        if (banded && idx > 0 && (idx % 2 === 1)) {
          r.shadingColor = bandColor;
        } else if (idx > 0) {
          r.shadingColor = "";
        }
      });

      await context.sync();
      setStatus("Style appliqué au tableau.");
    });
  } catch (err) {
    console.error(err);
    setStatus("Erreur lors de l'application du style : " + (err.message || err));
  }
}

async function validateTable() {
  try {
    await Word.run(async (context) => {
      const tables = context.document.body.tables;
      tables.load("items");
      await context.sync();

      if (tables.items.length === 0) {
        setStatus("Validation: échec — aucun tableau présent.");
        return;
      }
      const table = tables.items[0];
      table.getRow(0).load("shadingColor, font");
      table.rows.load("items");
      await context.sync();

      const headerShading = table.getRow(0).shadingColor;
      const rowsCount = table.rows.items.length;
      let bandedFound = false;
      for (let i = 1; i < table.rows.items.length; i++) {
        const row = table.rows.items[i];
        if (row.shadingColor && row.shadingColor !== "") {
          bandedFound = true; break;
        }
      }

      const ok = {
        tablePresent: true,
        rows: rowsCount,
        headerShaded: !!headerShading,
        banded: bandedFound
      };
      setStatus(`Validation: tableau présent (${ok.rows} lignes). En-tête stylée: ${ok.headerShaded}. Bandes détectées: ${ok.banded}.`);
    });
  } catch (err) {
    console.error(err);
    setStatus("Erreur lors de la validation : " + (err.message || err));
  }
}

async function clearDocument() {
  try {
    const doClear = confirm("Voulez-vous effacer tout le contenu du document ?");
    if (!doClear) { setStatus("Suppression annulée."); return; }
    await Word.run(async (context) => {
      context.document.body.clear();
      await context.sync();
      setStatus("Document effacé.");
    });
  } catch (err) {
    console.error(err);
    setStatus("Erreur lors de l'effacement : " + (err.message || err));
  }
}

export async function run() {
  return Word.run(async (context) => {
    /**
     * Insert your Word code here
     */

    const paragraph = context.document.body.insertParagraph("WAZA", Word.InsertLocation.end);

    paragraph.font.color = "blue";

    await context.sync();
  });
}
