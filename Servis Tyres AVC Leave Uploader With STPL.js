(async () => {

  if (location.href !== "https://hrms1.mydecibel.com/TMS/AVCD_BulkLeaves.aspx") {
    let p = document.createElement("div");
    p.textContent = "Not Allowed";
    p.style = "position:fixed;top:20px;left:50%;transform:translateX(-50%);background:red;color:#fff;padding:12px 25px;border-radius:8px;font-family:sans-serif;font-size:16px;z-index:999999;";
    document.body.appendChild(p);
    setTimeout(() => p.remove(), 2000);
    return;
  }

  let c = document.createElement("div");
  c.style = "position:fixed;top:5%;left:50%;transform:translateX(-50%);background:#fff;padding:40px 30px 30px 30px;z-index:999999;box-shadow:0 0 40px 10px rgba(0,0,0,0.8);border-radius:20px;text-align:center;font-family:sans-serif;min-width:500px;";

  c.innerHTML = `
    <div style="margin-bottom:15px;">
      <span style="font-weight:bold;font-size:26px;color:maroon;">AVC Leave Uploader</span>
    </div>

    <button id="dl" style="background:#000;color:#fff;padding:10px 14px;border:none;border-radius:8px;margin:8px 0;cursor:pointer;font-size:14px;">Download Template</button>

    <input type="file" id="up" accept=".csv" style="margin:8px;">

    <button id="ex" style="background:#217346;color:#fff;padding:10px 14px;border:none;border-radius:8px;margin:8px;cursor:pointer;font-size:14px;">Export to Excel</button>

    <div style="text-align:left;position:relative;margin-top:20px;">
      <button id="closeBtn" style="background:maroon;color:#fff;padding:7px 14px;border:none;border-radius:8px;cursor:pointer;position:absolute;top:0;left:0;">Close</button>
    </div>

    <div id="footerText" style="visibility:hidden;margin-top:20px;font-size:13px;color:#000;text-align:center;height:16px;">
      Prepared by Abdullah Shah and Mirza Laiq Ahmed
    </div>
  `;

  document.body.appendChild(c);

  let footer = document.getElementById("footerText");
  let timer = null;

  let hoverZone = document.createElement("div");
  hoverZone.style = "position:absolute;bottom:10px;right:10px;width:30px;height:30px;";
  c.appendChild(hoverZone);

  hoverZone.onmouseenter = () => {
    timer = setTimeout(() => {
      footer.style.visibility = "visible";
      footer.style.opacity = "1";
    }, 3000);
  };

  hoverZone.onmouseleave = () => {
    clearTimeout(timer);
    footer.style.visibility = "hidden";
  };

  closeBtn.onclick = () => c.remove();

  // ✅ CSV Template
  dl.onclick = () => {
    let a = document.createElement("a");
    a.href = "data:text/csv,Emp ID,Attendance Date,Leave Type,Remarks";
    a.download = "Leave_Template.csv";
    a.click();
  };

  // ✅ Export Table
  ex.onclick = () => {
    let t = document.querySelector("table");
    if (!t) return alert("❌ Table not found");

    let csv = [...t.rows]
      .map(r => [...r.cells]
        .map(c => `"${c.innerText.trim().replace(/"/g, '""')}"`)
        .join("\t"))
      .join("\n");

    let a = document.createElement("a");
    a.href = "data:application/vnd.ms-excel," + encodeURIComponent(csv);
    a.download = "Crystal_Report_Export.xls";
    a.click();
  };

  // ✅ CSV Upload (FIXED PARSER)
  up.onchange = e => {

    let f = e.target.files[0];
    if (!f) return;

    let reader = new FileReader();

    reader.onload = () => {

      let rows = reader.result
        .split("\n")
        .slice(1)
        .map(row => {
          let m = row.match(/(".*?"|[^",\n]+)(?=\s*,|\s*$)/g);
          return m ? m.map(v => v.replace(/^"|"$/g, "").trim()) : [];
        })
        .filter(r => r[0] && r[1] && r[2] && r[3]);

      let rejected = [];
      let i = 0;

      let search = document.querySelector("input.form-control.form-control-sm[type='search']");
      if (!search) return alert("❌ Search box not found");

      (function next() {

        if (i >= rows.length) {

          if (rejected.length) {
            let a = document.createElement("a");
            a.href = URL.createObjectURL(
              new Blob(
                ["EmployeeID,Date,LeaveCode,Remark\n" +
                  rejected.map(r => r.map(v => `"${v}"`).join(",")).join("\n")
                ],
                { type: "text/csv" }
              )
            );
            a.download = "Rejected_Leave_Entries.csv";
            a.click();
          }

          c.remove();
          return;
        }

        let [id, date, leave, remark] = rows[i++];

        search.value = id.trim();
        search.dispatchEvent(new Event("input"));

        setTimeout(() => {

          let found = false;

          document.querySelectorAll("table tbody tr").forEach(r => {

            let t = r.innerText.replace(/\s+/g, " ").trim();

            if (t.includes(id.trim()) && t.includes(date.trim())) {

              let cb = r.querySelector("input[type='checkbox']");
              cb && !cb.checked && cb.click();

              let s = r.querySelector("select");
              s && (s.value = leave,
                s.dispatchEvent(new Event("change", { bubbles: true })));

              let inputs = r.querySelectorAll("input[type='text']");
              inputs.length && (
                inputs[inputs.length - 1].value = remark,
                inputs[inputs.length - 1].dispatchEvent(new Event("input", { bubbles: true }))
              );

              found = true;
            }
          });

          !found && rejected.push([id, date, leave, "Emp not found"]);
          next();

        }, 400);

      })();
    };

    reader.readAsText(f);
  };

})();
