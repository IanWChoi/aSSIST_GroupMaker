let moduleName = "";
let studentNames = [];
let historyData = {}; // { moduleName: [ [group1], [group2], ... ] }
let uploadedHistoryWorkbook = null;
let lastScore = 0;
let exclusionPairs = [];

function nextStep(current, next = current + 1) {
  document.getElementById(`step${current}`).style.display = "none";
  document.getElementById(`step${next}`).style.display = "block";

  if (current === 1) {
    moduleName = document.getElementById("module").value.trim();
  }

  if (current === 2) {
    const raw = document.getElementById("names").value.trim();
    studentNames = raw.split("/").map((name) => name.trim()).filter(Boolean);

    const excludeRaw = document.getElementById("exclude").value.trim();
    exclusionPairs = excludeRaw
      .split("/")
      .map(pair => pair.split("-").map(name => name.trim()).sort().join("::"))
      .filter(Boolean);
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
  historyData = {};
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

/*  const homeButton = document.createElement("button");
  homeButton.innerText = "홈으로";
  homeButton.onclick = () => {
    location.reload();
  };
  document.getElementById("groupOutput").appendChild(homeButton);*/
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

    let invalid = false;
    for (const group of groups) {
      for (let i = 0; i < group.length; i++) {
        for (let j = i + 1; j < group.length; j++) {
          const key = [group[i], group[j]].sort().join("::");
          if (exclusionPairs.includes(key)) {
            invalid = true;
            break;
          }
        }
        if (invalid) break;
      }
      if (invalid) break;
    }
    if (invalid) continue;

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

  const moduleTitle = document.createElement("div");
  moduleTitle.id = "moduleDisplay";
  moduleTitle.innerText = moduleName;
  container.appendChild(moduleTitle);

  /*const scoreBox = document.createElement("div");
  scoreBox.id = "scoreBox";
  scoreBox.innerText = `중복 점수: ${lastScore}`;
  document.getElementById("result").insertBefore(scoreBox, container);*/

  groups.forEach((group, i) => {
    const div = document.createElement("div");
    div.innerHTML = `<strong>Group ${i + 1}</strong>: ${group.join(", ")}`;
    container.appendChild(div);
  });
}

function downloadHistory() {
  if (!uploadedHistoryWorkbook) uploadedHistoryWorkbook = XLSX.utils.book_new();

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

window.drawNetworkOnDemand = function () {
  const container = document.getElementById("network");
  container.innerHTML = ""; 

  // 중복 점수 표시
  const scoreOut = document.getElementById("scoreOutput");
  scoreOut.innerText = `중복 점수: ${lastScore}`;
  scoreOut.style.display = "block";

  drawNetwork(studentNames, Object.values(historyData).flat());
};

function drawNetwork(students, history) {
  const angleStep = (2 * Math.PI) / students.length;
  const radius = 50 + students.length * 10;

  const nodes = students.map((name, i) => ({
    id: name,
    label: name,
    x: radius * Math.cos(i * angleStep),
    y: radius * Math.sin(i * angleStep),
    fixed: true,
    font: {
      size: 16,
      vadjust: -5,
      color: "#2c3e50",
      face: "Segoe UI"
    },
    color: {
      background: "#ffffff",
      border: "#2d72d9"
    },
    shape: "dot",
    size: 10
  }));

  const edgeCount = {};
  history.forEach(group => {
    for (let i = 0; i < group.length; i++) {
      for (let j = i + 1; j < group.length; j++) {
        const key = [group[i], group[j]].sort().join("::");
        edgeCount[key] = (edgeCount[key] || 0) + 1;
      }
    }
  });

  const edges = Object.entries(edgeCount).map(([key, count]) => {
    const [a, b] = key.split("::");
    return {
      from: a,
      to: b,
      width: Math.min(1 + count, 5),
      color: {
        color: "#2d72d9",
        opacity: 0.4 + Math.min(count / 10, 0.5)
      }
    };
  });

  const container = document.getElementById("network");
  const data = { nodes, edges };
  const options = {
    layout: {
      improvedLayout: false
    },
    physics: false,
    edges: {
      smooth: {
        type: "continuous"
      }
    }
  };

  new vis.Network(container, data, options);
}

function captureResult() {
  const captureArea = document.getElementById('captureArea');
  const fileName = `${moduleName}_그룹.png`;
  
  html2canvas(captureArea, {
    backgroundColor: '#f7fafc',
    scale: 2,
    useCORS: true,
    allowTaint: false
  }).then(canvas => {
    const link = document.createElement('a');
    link.download = fileName;
    link.href = canvas.toDataURL('image/png');
    link.click();
  }).catch(error => {
    console.error('캡처 중 오류 발생:', error);
    alert('이미지 캡처에 실패했습니다.');
  });
}
