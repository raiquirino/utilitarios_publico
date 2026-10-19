    let currentFormat = null;

    function columnToLetter(n) {
      let s = "";
      while (n >= 0) {
        s = String.fromCharCode((n % 26) + 65) + s;
        n = Math.floor(n / 26) - 1;
      }
      return s;
    }

    function letterToColIndex(letter) {
      letter = letter.trim().toUpperCase();
      let col = 0;
      for (let i = 0; i < letter.length; i++) {
        col = col * 26 + (letter.charCodeAt(i) - 64);
      }
      return col - 1;
    }

    function excelDateToDDMMYYYY(v) {
      if (!isNaN(v) && v > 0 && v < 60000) {
        const base = new Date(Date.UTC(1899, 11, 30));
        const d = new Date(base.getTime() + v * 86400000);
        return `${String(d.getUTCDate()).padStart(2, "0")}/${String(d.getUTCMonth() + 1).padStart(2, "0")}/${d.getUTCFullYear()}`;
      }
      return v;
    }

    function formatValue(v) {
      const num = parseFloat(v);
      return isNaN(num) ? v : num.toLocaleString("pt-BR", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
    }

    function setFormatButtons(activeId) {
      ["btnDate", "btnValue", "btnOriginal", "btnDesativar"].forEach(id => {
        document.getElementById(id).classList.toggle("active", id === activeId);
      });
    }

    document.getElementById("btnDate").onclick = () => { currentFormat = "date"; setFormatButtons("btnDate"); };
    document.getElementById("btnValue").onclick = () => { currentFormat = "value"; setFormatButtons("btnValue"); };
    document.getElementById("btnOriginal").onclick = () => { currentFormat = "original"; setFormatButtons("btnOriginal"); };
    document.getElementById("btnDesativar").onclick = () => { currentFormat = null; setFormatButtons(null); };

    document.getElementById("inputExcel").addEventListener("click", function () { this.value = ""; });

    function aplicarMacro(ws) {
      ws[3][0]  = ws[3][2];
      ws[3][4]  = ws[3][3];
      ws[3][7]  = ws[3][8];
      ws[3][11] = ws[3][13];
      ws[3][15] = ws[3][14];
      ws[3][19] = ws[3][18];
      ws.splice(0, 3);

      function removerColunas(inicio, qtd) {
        ws.forEach(linha => linha.splice(inicio, qtd));
      }
      removerColunas(1, 3);
      removerColunas(2, 2);
      removerColunas(3, 3);
      removerColunas(4, 3);
      removerColunas(5, 3);
      removerColunas(6, 6);

      ws = ws.filter((linha, i) => i === 0 || linha[3] !== "");
      ws.forEach((linha, i) => {
        if (i !== 0) {
          if (linha[4] === 0) linha[4] = "";
          if (linha[5] === 0) linha[5] = "";
        }
      });
      return ws;
    }

    function processarArquivo(file) {
      if (!file) return;

      const reader = new FileReader();
      reader.onload = function (event) {
        const workbook = XLSX.read(new Uint8Array(event.target.result), { type: "array" });
        let html = "";

        workbook.SheetNames.forEach(name => {
          const sheet = workbook.Sheets[name];
          let rows = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true, defval: "" });

          if (document.getElementById("usarMacro1")?.checked) {
            rows = aplicarMacro(rows);
          }

          const maxCols = Math.max(...rows.map(r => r.length), 1);

          let table = "<table><tr>";
          for (let c = 0; c < maxCols; c++) {
            table += `<th data-col="${c}">${columnToLetter(c)}</th>`;
          }
          table += "</tr>";

          rows.forEach(row => {
            table += "<tr>";
            for (let c = 0; c < maxCols; c++) {
              const val = row[c] ?? "";
              const safeVal = String(val).replace(/"/g, "&quot;");
              const safeTitle = String(val).replace(/"/g, "&quot;").replace(/</g, "&lt;");
              table += `<td data-col="${c}" data-original="${safeVal}" title="${safeTitle}">${val}</td>`;
            }
            table += "</tr>";
          });
          table += "</table>";
          html += table;
        });

        document.getElementById("output").innerHTML = html;

        document.querySelectorAll("#output td, #output th").forEach(cell => {
          cell.addEventListener("click", function () {
            if (!currentFormat) return;
            const col = parseInt(this.getAttribute("data-col"));
            const table = this.closest("table");
            table.querySelectorAll(`td[data-col="${col}"]`).forEach(td => {
              const original = td.getAttribute("data-original");
              if (currentFormat === "date") {
                td.textContent = excelDateToDDMMYYYY(original);
                td.classList.remove("num");
              } else if (currentFormat === "value") {
                td.textContent = formatValue(original);
                td.classList.add("num");
              } else {
                td.textContent = original;
                td.classList.remove("num");
              }
            });
            atualizarTotais();
          });
        });

        atualizarTotais();
      };
      reader.readAsArrayBuffer(file);
    }

    document.getElementById("inputExcel").addEventListener("change", function (e) {
      processarArquivo(e.target.files[0]);
    });

    // Reprocessa o arquivo já carregado quando a checkbox da Macro é alterada
    document.getElementById("usarMacro1").addEventListener("change", function () {
      const file = document.getElementById("inputExcel").files[0];
      if (file) {
        processarArquivo(file);
      }
    });

    document.getElementById("btnConciliar").addEventListener("click", () => {
      const colA = letterToColIndex(document.getElementById("col1").value);
      const colB = letterToColIndex(document.getElementById("col2").value);
      const colR = letterToColIndex(document.getElementById("colR").value);

      if (colA < 0 || colB < 0 || colR < 0) {
        alert("Informe colunas válidas (ex: A, B, AA...)");
        return;
      }

      document.querySelectorAll("#output table").forEach(table => {
        const rows = table.querySelectorAll("tr");
        let numCols = rows[0].querySelectorAll("th").length;

        if (colR >= numCols) {
          for (let r = 0; r < rows.length; r++) {
            if (r === 0) rows[r].insertAdjacentHTML("beforeend", `<th data-col="${colR}">${columnToLetter(colR)}</th>`);
            else if (r === 1) rows[r].insertAdjacentHTML("beforeend", `<td data-col="${colR}" data-original="Conciliação">Conciliação</td>`);
            else rows[r].insertAdjacentHTML("beforeend", `<td data-col="${colR}" data-original=""></td>`);
          }
        }

        for (let r = 2; r < rows.length; r++) {
          const res = rows[r].querySelector(`td[data-col="${colR}"]`);
          if (res) { res.textContent = ""; res.setAttribute("data-original", ""); }
        }

        const usadosB = new Set();
        for (let r1 = 2; r1 < rows.length; r1++) {
          const cel1 = rows[r1].querySelector(`td[data-col="${colA}"]`);
          if (!cel1) continue;
          const v1 = cel1.textContent.trim();
          if (!v1) continue;

          for (let r2 = 2; r2 < rows.length; r2++) {
            if (usadosB.has(r2)) continue;
            const cel2 = rows[r2].querySelector(`td[data-col="${colB}"]`);
            if (!cel2) continue;
            const v2 = cel2.textContent.trim();
            if (!v2) continue;

            if (v1 === v2) {
              usadosB.add(r2);
              const res1 = rows[r1].querySelector(`td[data-col="${colR}"]`);
              const res2 = rows[r2].querySelector(`td[data-col="${colR}"]`);
              if (res1) { res1.textContent = "Sim"; res1.setAttribute("data-original", "Sim"); }
              if (res2) { res2.textContent = "Sim"; res2.setAttribute("data-original", "Sim"); }
              break;
            }
          }
        }

        [colA, colB].forEach(col => {
          table.querySelectorAll(`td[data-col="${col}"]`).forEach(td => {
            td.textContent = formatValue(td.getAttribute("data-original"));
            td.classList.add("num");
          });
        });
      });

      atualizarTotais();
    });

    function filtrarTabela(tipo) {
      const colR = letterToColIndex(document.getElementById("colR").value);
      if (colR < 0) return;

      document.querySelectorAll("#output table").forEach(table => {
        const linhas = table.querySelectorAll("tr");
        for (let i = 2; i < linhas.length; i++) {
          const cel = linhas[i].querySelector(`td[data-col="${colR}"]`);
          const valor = cel ? cel.textContent.trim().toLowerCase() : "";
          if (tipo === "conciliados") linhas[i].style.display = (valor === "sim") ? "" : "none";
          else if (tipo === "nao") linhas[i].style.display = (valor === "") ? "" : "none";
          else linhas[i].style.display = "";
        }
      });
      atualizarTotais();
    }

    document.getElementById("btnConciliados").onclick = () => filtrarTabela("conciliados");
    document.getElementById("btnNaoConciliados").onclick = () => filtrarTabela("nao");
    document.getElementById("btnTodos").onclick = () => filtrarTabela("todos");

    function atualizarTotais() {
      const colA = letterToColIndex(document.getElementById("col1").value);
      const colB = letterToColIndex(document.getElementById("col2").value);
      let totalA = 0, totalB = 0;
      const letraA = document.getElementById("col1").value.toUpperCase() || "?";
      const letraB = document.getElementById("col2").value.toUpperCase() || "?";

      document.querySelectorAll("#output table").forEach(table => {
        table.querySelectorAll("tr").forEach((linha, i) => {
          if (i < 2 || linha.style.display === "none") return;
          const celA = linha.querySelector(`td[data-col="${colA}"]`);
          const celB = linha.querySelector(`td[data-col="${colB}"]`);
          if (celA) {
            const n = parseFloat(celA.textContent.trim().replace(/\./g, "").replace(",", "."));
            if (!isNaN(n)) totalA += n;
          }
          if (celB) {
            const n = parseFloat(celB.textContent.trim().replace(/\./g, "").replace(",", "."));
            if (!isNaN(n)) totalB += n;
          }
        });
      });

      const el = document.getElementById("totais");
      el.innerHTML = `Total Coluna <strong>${letraB}</strong>: ${totalB.toLocaleString("pt-BR", { minimumFractionDigits: 2, maximumFractionDigits: 2 })}
        &nbsp;&nbsp;|&nbsp;&nbsp;
        Total Coluna <strong>${letraA}</strong>: ${totalA.toLocaleString("pt-BR", { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;
      el.classList.add("visible");
    }

    document.getElementById("btnSalvar").addEventListener("click", () => {
      const tables = document.querySelectorAll("#output table");
      if (tables.length === 0) { alert("Nenhuma tabela para salvar."); return; }

      tables.forEach(table => {
        const workbook = XLSX.utils.book_new();
        let data = [];
        const linhas = table.querySelectorAll("tr");

        for (let i = 1; i < linhas.length; i++) {
          if (i > 1 && linhas[i].style.display === "none") continue;
          const cells = linhas[i].querySelectorAll("th, td");
          let row = [];
          cells.forEach(cell => {
            let val = cell.textContent;
            const colIndex = parseInt(cell.getAttribute("data-col"));
            const col1Index = letterToColIndex(document.getElementById("col1").value);
            const col2Index = letterToColIndex(document.getElementById("col2").value);
            if (i > 1 && (colIndex === col1Index || colIndex === col2Index)) {
              val = parseFloat(val.replace(/\./g, "").replace(",", "."));
              if (isNaN(val)) val = "";
            }
            row.push(val);
          });
          data.push(row);
        }

        const ws = XLSX.utils.aoa_to_sheet(data);
        XLSX.utils.book_append_sheet(workbook, ws, "Sheet1");
        const nome = (document.getElementById("inputExcel").files[0]?.name || "planilha").replace(/\.[^/.]+$/, "") + "_conciliado.xlsx";
        XLSX.writeFile(workbook, nome);
      });
    });

    document.querySelector("#output").addEventListener("click", function (e) {
      if (!e.ctrlKey) return;
      if (e.target.tagName === "TD") {
        e.preventDefault();
        e.stopPropagation();
        e.target.parentElement.classList.toggle("selected-row");
      }
    });

    document.querySelector("#output").addEventListener("dblclick", function (e) {
      if (e.target.tagName === "TD") {
        e.preventDefault();
        e.stopPropagation();
        const texto = e.target.textContent.trim();
        if (texto) {
          navigator.clipboard.writeText(texto).then(() => {
            const msg = document.getElementById("copiado-msg");
            msg.textContent = `Copiado: ${texto}`;
            msg.style.display = "block";
            setTimeout(() => msg.style.display = "none", 1400);
          });
        }
      }
    });