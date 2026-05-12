function buildInlineNoteCell(td, noteValue, onSave) {
  td.textContent = "";
  const hasNote = noteValue && noteValue !== "—" && noteValue !== "";

  if (hasNote) {
    const span = document.createElement("span");
    span.className = "note-text";
    span.textContent = noteValue;
    td.appendChild(span);
  }

  const btn = document.createElement("button");
  btn.type = "button";
  btn.className = "note-add-btn";
  btn.textContent = hasNote ? "✎" : "+ Add note";
  btn.addEventListener("click", (e) => {
    e.stopPropagation();
    startInlineEdit(td, noteValue, onSave);
  });
  td.appendChild(btn);
}

function startInlineEdit(td, noteValue, onSave) {
  if (td.querySelector("input")) return;
  const currentNote = noteValue && noteValue !== "—" ? noteValue : "";

  td.textContent = "";
  const input = document.createElement("input");
  input.type = "text";
  input.value = currentNote;
  input.className = "note-inline-input";
  input.placeholder = "Add a note…";
  input.maxLength = 200;
  td.appendChild(input);
  input.focus();
  input.select();

  let committed = false;

  function commit() {
    if (committed) return;
    committed = true;
    const note = input.value.trim();
    buildInlineNoteCell(td, note, onSave);
    onSave(note).catch(() => buildInlineNoteCell(td, currentNote, onSave));
  }

  function cancel() {
    if (committed) return;
    committed = true;
    buildInlineNoteCell(td, currentNote, onSave);
  }

  input.addEventListener("blur", commit);
  input.addEventListener("keydown", (e) => {
    if (e.key === "Enter") {
      e.preventDefault();
      input.removeEventListener("blur", commit);
      commit();
    } else if (e.key === "Escape") {
      e.preventDefault();
      input.removeEventListener("blur", commit);
      cancel();
    }
  });
}

function initNoteCells(tbodyId, apiEndpoint) {
  const tbody = document.getElementById(tbodyId);
  if (!tbody) return;
  tbody.querySelectorAll("tr[data-rowkey]").forEach((tr) => {
    const noteTd = tr.querySelector(".note-cell");
    if (!noteTd) return;
    const rowkey = tr.dataset.rowkey;
    let noteValue = tr.dataset.note || "";
    buildInlineNoteCell(noteTd, noteValue, (note) => {
      noteValue = note;
      return fetch(apiEndpoint, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ rowkey, note }),
      }).then((r) => (r.ok ? r.json() : Promise.reject()));
    });
  });
}
