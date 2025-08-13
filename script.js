let moduleName = "";
let studentNames = [];
let historyData = {}; // { moduleName: [ [group1], [group2], ... ] }
let uploadedHistoryWorkbook = null;
let lastScore = 0;
let exclusionPairs = [];

function nextStep(current, next = current + 1) {
  if (current === 1) {
    moduleName = document.getElementById("module").value.trim();
  }

  if (current === 2) {
    const raw = document.getElementById("names").value.trim();
    const names = raw.split("/").map((name) => name.trim()).filter(Boolean);

    // 학생 수 제한 (200명)
    if (names.length > 200) {
      alert('학생 수는 200명을 초과할 수 없습니다.');
      return;
    }

    const uniqueNames = new Set(names);
    if (uniqueNames.size < names.length) {
      const counts = {};
      names.forEach(name => { counts[name] = (counts[name] || 0) + 1; });
      const duplicates = Object.keys(counts).filter(name => counts[name] > 1);
      alert(`중복된 학생 이름이 있습니다: ${duplicates.join(', ')}
확인 후 다시 시도해주세요.`);
      return;
    }
    
    studentNames = names;

    const excludeRaw = document.getElementById("exclude").value.trim();
    exclusionPairs = excludeRaw
      .split("/")
      .map(pair => pair.split("-").map(name => name.trim()).sort().join("::"))
      .filter(Boolean);
  }

  document.getElementById(`step${current}`).style.display = "none";
  document.getElementById(`step${next}`).style.display = "block";
}

function handleHistoryUpload() {
  document.getElementById("historyFile").style.display = "block";
  document.getElementById("historyFile").addEventListener("change", (e) => {
    const file = e.target.files[0];
    
    // 파일 크기 제한 (5MB)
    if (file.size > 5 * 1024 * 1024) {
      alert('파일 크기는 5MB를 초과할 수 없습니다.');
      e.target.value = '';
      return;
    }
    
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
  
  if (!result.groups) {
    alert("제외 조합 조건에 맞는 조 편성을 찾을 수 없습니다. 제외 조합을 줄이거나 학생 수를 조정해주세요.");
    return;
  }

  const groups = result.groups;
  lastScore = result.score;

  if (!historyData[moduleName]) {
    historyData[moduleName] = [];
  }
  // 이전 라운드의 결과를 덮어쓰지 않고, 새 결과로 대체합니다.
  historyData[moduleName] = groups;

  displayGroups(groups);
  document.getElementById("step4").style.display = "none";
  document.getElementById("result").style.display = "block";
  
  const scoreOut = document.getElementById("scoreOutput");
  scoreOut.textContent = `중복 점수: ${lastScore}`;
  scoreOut.style.display = "block";
}

function calculateScore(groups, history) {
  const pairCounts = {};
  history.forEach((group) => {
    for (let i = 0; i < group.length; i++) {
      for (let j = i + 1; j < group.length; j++) {
        const key = [group[i], group[j]].sort().join("::");
        pairCounts[key] = (pairCounts[key] || 0) + 1;
      }
    }
  });

  let score = 0;
  for (const group of groups) {
    for (let i = 0; i < group.length; i++) {
      for (let j = i + 1; j < group.length; j++) {
        const key = [group[i], group[j]].sort().join("::");
        score += pairCounts[key] || 0;
      }
    }
  }
  return score;
}

function generateGroups(students, numGroups, history) {
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

    const score = calculateScore(groups, history);

    if (score < lowestScore) {
      lowestScore = score;
      best = groups;
    }
  }

  return { groups: best, score: lowestScore };
}

function displayGroups(groups) {
  const container = document.getElementById("groupOutput");
  container.replaceChildren();

  // HTML에 이미 있는 moduleDisplay 사용
  const moduleTitle = document.getElementById("moduleDisplay");
  moduleTitle.textContent = moduleName;

  groups.forEach((group, i) => {
    const div = document.createElement("div");
    const strong = document.createElement("strong");
    strong.textContent = `Group ${i + 1}`;
    div.appendChild(strong);
    
    const ul = document.createElement("ul");
    ul.className = "group-list";
    ul.dataset.groupIndex = i;

    group.forEach(student => {
      const li = document.createElement("li");
      li.className = "student-item";
      li.textContent = student;
      li.dataset.studentName = student;
      ul.appendChild(li);
    });

    div.appendChild(ul);
    container.appendChild(div);

    new Sortable(ul, {
      group: 'shared',
      animation: 150,
      onEnd: (evt) => {
        const studentName = evt.item.dataset.studentName;
        const fromGroupIndex = parseInt(evt.from.dataset.groupIndex);
        const toGroupIndex = parseInt(evt.to.dataset.groupIndex);
        const oldIndex = evt.oldDraggableIndex;
        const newIndex = evt.newDraggableIndex;

        // Update internal data structure
        const currentGroups = historyData[moduleName];
        
        // Remove from old group
        currentGroups[fromGroupIndex].splice(oldIndex, 1);
        
        // Add to new group
        currentGroups[toGroupIndex].splice(newIndex, 0, studentName);

        // Recalculate and update score
        const allHistory = Object.values(historyData).filter(h => h !== currentGroups).flat();
        lastScore = calculateScore(currentGroups, allHistory);
        
        const scoreOut = document.getElementById("scoreOutput");
        scoreOut.textContent = `중복 점수: ${lastScore}`;
      }
    });
  });
}

function sanitizeForExcel(value) {
  // 엑셀 수식 인젝션 방지: =, +, -, @ 로 시작하는 값에 아포스트로피 추가
  if (typeof value === 'string' && /^[=+\-@]/.test(value)) {
    return "'" + value;
  }
  return value;
}

function sanitizeModuleName(name) {
  // 파일명에 사용할 수 없는 문자 제거 및 길이 제한
  return name
    .replace(/[<>:"/\\|?*\x00-\x1f]/g, '') // 제어 문자 및 특수 문자 제거
    .trim()
    .substring(0, 50); // 최대 50자로 제한
}

function downloadHistory() {
  const wb = XLSX.utils.book_new();

  // Add current module's modified data
  const currentModuleData = [];
  const currentGroups = historyData[moduleName];
  currentGroups.forEach((group, idx) => {
    const sanitizedGroup = group.map(student => sanitizeForExcel(student));
    currentModuleData.push([idx + 1, ...sanitizedGroup]);
  });
  const ws = XLSX.utils.aoa_to_sheet([["조번호", "학생1", "학생2", "..."]].concat(currentModuleData));
  XLSX.utils.book_append_sheet(wb, ws, sanitizeModuleName(moduleName));

  // Add other sheets from original workbook if they exist
  if (uploadedHistoryWorkbook) {
    uploadedHistoryWorkbook.SheetNames.forEach(sheetName => {
      if (sheetName !== moduleName) {
        const originalWs = uploadedHistoryWorkbook.Sheets[sheetName];
        XLSX.utils.book_append_sheet(wb, originalWs, sheetName);
      }
    });
  }
  
  XLSX.writeFile(wb, `history_${sanitizeModuleName(moduleName)}_updated.xlsx`);
}

window.drawNetworkOnDemand = function () {
  const container = document.getElementById("network");
  container.replaceChildren(); 

  const scoreOut = document.getElementById("scoreOutput");
  scoreOut.textContent = `중복 점수: ${lastScore}`;
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

// DOM이 로드된 후 이벤트 리스너 등록
document.addEventListener('DOMContentLoaded', function() {
  // 제목 클릭시 페이지 새로고침
  document.getElementById('title').addEventListener('click', function() {
    location.reload();
  });
  
  // 단계별 다음 버튼
  document.getElementById('step1-next').addEventListener('click', function() {
    nextStep(1);
  });
  
  document.getElementById('step2-next').addEventListener('click', function() {
    nextStep(2);
  });
  
  // 히스토리 관련 버튼
  document.getElementById('upload-history').addEventListener('click', function() {
    handleHistoryUpload();
  });
  
  document.getElementById('skip-history').addEventListener('click', function() {
    skipHistory();
  });
  
  // 조편성 시작 버튼
  document.getElementById('start-grouping').addEventListener('click', function() {
    runGrouping();
  });
  
  // 결과 화면 버튼들
  document.getElementById('home-btn').addEventListener('click', function() {
    location.reload();
  });
  
  document.getElementById('download-btn').addEventListener('click', function() {
    downloadHistory();
  });
  
  document.getElementById('capture-btn').addEventListener('click', function() {
    captureResult();
  });
  
  document.getElementById('visualize-btn').addEventListener('click', function() {
    drawNetworkOnDemand();
  });
});