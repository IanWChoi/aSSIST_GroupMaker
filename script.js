let moduleName = "";
let studentNames = [];
let historyData = {}; // { moduleName: [ [group1], [group2], ... ] }
let uploadedHistoryWorkbook = null;
let lastScore = 0;

function nextStep(step) {
  document.getElementById(`step${step}`).style.display = "none";
  document.getElementById(`step${step + 1}`).style.display = "block";

  if (step === 1) {
    moduleName = document.getElementById("module").value.trim();
  }

  if (step === 2) {
    const raw = document.getElementById("names").value.trim();
    studentNames = raw.split("/").map((name) => name.trim()).filter(Boolean);
  }
}

function handleHistoryUpload() {
  document.getElementById("historyFile").style.display = "block";
  document.getElementById("historyFile").addEventListener("change", (e) => {
    const file = e.target.files[0];
    const reader = new FileReader();
    reader.onload = (evt) => {
      const data = new Uint8Array(evt.target.result);
      uploadedHistoryWorkbook = XLSX.read(data, { type: "array" });
      uploadedHistoryWorkbook.SheetNames.forEach((sheetName) => {
        const sheet = uploadedHistoryWorkbook.Sheets[sheetName];
        const json = XLSX.utils.sheet_to_json(sheet, { header: 1 }).slice(1);
        if (!historyData[sheetName]) historyData[sheetName] = [];
        json.forEach((row) => {
          const group = row.slice(1).filter(Boolean);
          if (group.length > 0) {
            historyData[sheetName].push(group);
          }
        });
      });
      nextStep(3);
    };
    reader.readAsArrayBuffer(file);
  });
}

function skipHistory() {
  uploadedHistoryWorkbook = XLSX.utils.book_new();
  historyData = {}; // 초기화
  nextStep(3);
}

function runGrouping() {
  const numGroups = parseInt(document.getElementById("numGroups").value);
  if (!numGroups || numGroups <= 0 || studentNames.length < numGroups) {
    alert("유효한 조 개수를 입력해주세요.");
    return;
  }

  const allHistory = Object.values(historyData).flat();
  const result = generateGroups(studentNames, numGroups, allHistory);
  const groups = result.groups;
  lastScore = result.score;

  if (!historyData[moduleName]) historyData[moduleName] = [];
  historyData[moduleName].push(...groups);

  displayGroups(groups);
  document.getElementById("step4").style.display = "none";
  document.getElementById("result").style.display = "block";

  const homeButton = document.createElement("button");
  homeButton.innerText = "홈으로";
  homeButton.onclick = () => {
    location.reload();
  };
  document.getElementById("groupOutput").appendChild(homeButton);
}

function generateGroups(students, numGroups, history) {
  const pairCounts = {};
  history.forEach((group) => {
    for (let i = 0; i < group.length; i++) {
      for (let j = i + 1; j < group.length; j++) {
        const key = [group[i], group[j]].sort().join("::");
        pairCounts[key] = (pairCounts[key] || 0) + 1;
      }
    }
  });

  let best = null;
  let lowestScore = Infinity;

  for (let t = 0; t < 1000; t++) {
    const shuffled = [...students].sort(() => Math.random() - 0.5);
    const groups = Array.from({ length: numGroups }, () => []);
    for (let i = 0; i < shuffled.length; i++) {
      groups[i % numGroups].push(shuffled[i]);
    }

    let score = 0;
    for (const group of groups) {
      for (let i = 0; i < group.length; i++) {
        for (let j = i + 1; j < group.length; j++) {
          const key = [group[i], group[j]].sort().join("::");
          score += pairCounts[key] || 0;
        }
      }
    }

    if (score < lowestScore) {
      lowestScore = score;
      best = groups;
    }
  }

  return { groups: best, score: lowestScore };
}

function displayGroups(groups) {
  const container = document.getElementById("groupOutput");
  container.innerHTML = "";

  const scoreDiv = document.createElement("div");
  scoreDiv.innerHTML = `<p><strong>🔁 중복 점수:</strong> ${lastScore}</p>`;
  container.appendChild(scoreDiv);

  groups.forEach((group, i) => {
    const div = document.createElement("div");
    div.innerHTML = `<strong>Group ${i + 1}</strong>: ${group.join(", ")}`;
    container.appendChild(div);
  });
}

function downloadHistory() {
  if (!uploadedHistoryWorkbook) uploadedHistoryWorkbook = XLSX.utils.book_new();

  // 현재 모듈 시트 제거 (있다면)
  const idx = uploadedHistoryWorkbook.SheetNames.indexOf(moduleName);
  if (idx !== -1) uploadedHistoryWorkbook.SheetNames.splice(idx, 1);

  const data = [];
  const history = historyData[moduleName];
  history.forEach((group, idx) => {
    data.push([idx + 1, ...group]);
  });
  const ws = XLSX.utils.aoa_to_sheet([["조번호", "학생1", "학생2", "..."]].concat(data));
  XLSX.utils.book_append_sheet(uploadedHistoryWorkbook, ws, moduleName);

  XLSX.writeFile(uploadedHistoryWorkbook, `history_${moduleName}.xlsx`);
}
