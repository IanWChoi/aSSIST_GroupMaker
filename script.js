let moduleName = "";
let studentNames = [];
let studentData = []; // 고급 모드용: [{name, gender, job}, ...]
let historyData = {}; // { moduleName: [ [group1], [group2], ... ] }
let uploadedHistoryWorkbook = null;
let lastScore = 0;
let exclusionPairs = [];
let currentMode = 'basic'; // 'basic' or 'advanced'
let balanceOptions = {
  gender: true,
  job: true
};
let priorityOrder = ['history', 'gender', 'job']; // 기본 우선순위

function nextStep(current, next = current + 1) {
  if (current === 1) {
    moduleName = document.getElementById("module").value.trim();
    
    // 모듈명 31자 제한 체크 (Excel 시트명 제한)
    if (moduleName.length > 31) {
      alert('모듈 이름은 31자 이하로 입력해주세요.\n(Excel 시트명 제한)');
      return;
    }
    
    // 첫 번째 단계를 벗어나면 공지 숨기기
    document.querySelector(".notice-text").style.display = "none";
  }

  if (current === 2) {
    if (currentMode === 'basic') {
      // 기본 모드 처리
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
      studentData = []; // 고급 모드 데이터 초기화

      const excludeRaw = document.getElementById("exclude").value.trim();
      exclusionPairs = excludeRaw
        .split("/")
        .map(pair => pair.split("-").map(name => name.trim()).sort().join("::"))
        .filter(Boolean);
    } else {
      // 고급 모드 처리
      const dataRaw = document.getElementById("student-data").value.trim();
      if (!dataRaw) {
        alert('학생 정보를 입력해주세요.');
        return;
      }

      const lines = dataRaw.split('\n').filter(line => line.trim());
      const parsedData = [];
      const names = [];

      for (const line of lines) {
        const parts = line.split('/').map(part => part.trim());
        if (parts.length < 1) {
          alert(`형식 오류: "${line}"\n최소 이름은 있어야 합니다.`);
          return;
        }
        
        const name = parts[0];
        const gender = parts[1] || '';
        const job = parts[2] || '';
        
        if (!name) {
          alert(`이름이 없습니다: "${line}"`);
          return;
        }
        
        let studentInfo = { name };
        
        if (gender) {
          if (!['남', '여', 'M', 'F', '남성', '여성'].includes(gender)) {
            alert(`성별은 '남' 또는 '여'로 입력해주세요: "${line}"`);
            return;
          }
          
          studentInfo.gender = gender === '남' || gender === '남성' || gender === 'M' ? '남' : '여';
        }
        
        if (job) {
          studentInfo.job = job;
        }
        
        parsedData.push(studentInfo);
        names.push(name);
      }

      // 중복 체크
      const uniqueNames = new Set(names);
      if (uniqueNames.size < names.length) {
        const counts = {};
        names.forEach(name => { counts[name] = (counts[name] || 0) + 1; });
        const duplicates = Object.keys(counts).filter(name => counts[name] > 1);
        alert(`중복된 학생 이름이 있습니다: ${duplicates.join(', ')}\n확인 후 다시 시도해주세요.`);
        return;
      }

      if (parsedData.length > 200) {
        alert('학생 수는 200명을 초과할 수 없습니다.');
        return;
      }

      studentData = parsedData;
      studentNames = names;
      
      // 고급 모드 옵션 가져오기
      balanceOptions.gender = document.getElementById('balance-gender').checked;
      balanceOptions.job = document.getElementById('balance-job').checked;

      const excludeRaw = document.getElementById("exclude-advanced").value.trim();
      exclusionPairs = excludeRaw
        .split("/")
        .map(pair => pair.split("-").map(name => name.trim()).sort().join("::"))
        .filter(Boolean);
    }
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
  let result;
  
  if (currentMode === 'advanced' && studentData.length > 0) {
    // 고급 모드: 성비/직군 분산 알고리즘 사용
    result = generateAdvancedGroups(studentData, numGroups, allHistory);
  } else {
    // 기본 모드: 기존 알고리즘 사용
    result = generateGroups(studentNames, numGroups, allHistory);
  }
  
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
  if (currentMode === 'advanced') {
    scoreOut.innerHTML = `중복 점수: ${lastScore}<br>성비/직군 균형도: ${result.balanceScore || 'N/A'}`;
  } else {
    scoreOut.textContent = `중복 점수: ${lastScore}`;
  }
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

// 고급 모드: 성비/직군 분산 알고리즘
function generateAdvancedGroups(students, numGroups, history) {
  let best = null;
  let lowestScore = Infinity;
  let bestBalanceScore = Infinity;

  for (let t = 0; t < 1000; t++) {
    const shuffled = [...students].sort(() => Math.random() - 0.5);
    const groups = Array.from({ length: numGroups }, () => []);
    
    // 성비와 직군을 고려한 분배
    if (balanceOptions.gender || balanceOptions.job) {
      // 학생을 성별로 분류
      const categorized = {
        남: shuffled.filter(s => s.gender === '남'),
        여: shuffled.filter(s => s.gender === '여'),
        미정: shuffled.filter(s => !s.gender)
      };
      
      // 각 그룹에 균등하게 분배
      let groupIndex = 0;
      
      // 성별 균등 분배
      if (balanceOptions.gender) {
        for (const gender of ['남', '여', '미정']) {
          categorized[gender].forEach(student => {
            groups[groupIndex % numGroups].push(student.name);
            groupIndex++;
          });
        }
      } else {
        // 성별 고려 없이 분배
        shuffled.forEach((student, i) => {
          groups[i % numGroups].push(student.name);
        });
      }
    } else {
      // 옵션이 없으면 기본 분배
      shuffled.forEach((student, i) => {
        groups[i % numGroups].push(student.name);
      });
    }

    // 제외 조합 체크
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
    const balanceScore = calculateBalanceScore(groups, students);
    const weights = getPriorityWeights();

    // 우선순위 기반 종합 점수 계산
    let totalScore = 0;
    
    // history (중복 최소화)
    if (weights.history) {
      totalScore += score * weights.history;
    }
    
    // gender와 job 점수를 분리하여 계산
    const genderScore = balanceOptions.gender ? calculateGenderBalanceScore(groups, students) : 0;
    const jobScore = balanceOptions.job ? calculateJobBalanceScore(groups, students) : 0;
    
    if (weights.gender && balanceOptions.gender) {
      totalScore += genderScore * weights.gender;
    }
    
    if (weights.job && balanceOptions.job) {
      totalScore += jobScore * weights.job;
    }

    if (totalScore < lowestScore) {
      lowestScore = totalScore;
      bestBalanceScore = genderScore + jobScore;
      best = groups;
    }
  }

  return { 
    groups: best, 
    score: lowestScore,
    balanceScore: bestBalanceScore.toFixed(2)
  };
}

// 균형도 점수 계산 (기존 함수 유지 - 호환성용)
function calculateBalanceScore(groups, students) {
  return calculateGenderBalanceScore(groups, students) + calculateJobBalanceScore(groups, students);
}

// 성별 균형 점수 계산
function calculateGenderBalanceScore(groups, students) {
  let score = 0;
  
  groups.forEach(group => {
    const genderCounts = { 남: 0, 여: 0, 미정: 0 };
    group.forEach(name => {
      const student = students.find(s => s.name === name);
      if (student && student.gender) {
        genderCounts[student.gender]++;
      } else {
        genderCounts.미정++;
      }
    });
    const genderDiff = Math.abs(genderCounts.남 - genderCounts.여);
    score += genderDiff;
  });
  
  return score;
}

// 직군 균형 점수 계산
function calculateJobBalanceScore(groups, students) {
  let score = 0;
  
  groups.forEach(group => {
    const jobCounts = {};
    group.forEach(name => {
      const student = students.find(s => s.name === name);
      if (student && student.job) {
        jobCounts[student.job] = (jobCounts[student.job] || 0) + 1;
      } else {
        jobCounts['미정'] = (jobCounts['미정'] || 0) + 1;
      }
    });
    
    // 가장 많은 직군의 수를 기준으로 다양성 점수 계산
    const counts = Object.values(jobCounts);
    if (counts.length > 0) {
      const maxCount = Math.max(...counts);
      const diversity = group.length / Object.keys(jobCounts).length;
      score += diversity;
    }
  });
  
  return score;
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
      
      // 고급 모드일 때 성별에 따른 색상 클래스 추가
      if (currentMode === 'advanced' && studentData.length > 0) {
        const studentInfo = studentData.find(s => s.name === student);
        if (studentInfo && studentInfo.gender) {
          if (studentInfo.gender === '남') {
            li.classList.add('gender-male');
          } else if (studentInfo.gender === '여') {
            li.classList.add('gender-female');
          }
        } else {
          li.classList.add('gender-unknown');
        }
      }
      
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

// 샘플 엑셀 다운로드 함수
function downloadSampleExcel() {
  const wb = XLSX.utils.book_new();
  
  // 샘플 데이터
  const sampleData = [
    ['이름', '성별', '직군'],
    ['김철수', '남', '개발'],
    ['이영희', '여', '기획'],
    ['박민수', '남', '디자인'],
    ['정수진', '여', '개발'],
    ['최동현', '남', '마케팅'],
    ['김서연', '여', 'QA'],
    ['이준호', '남', '기획'],
    ['박지은', '여', '디자인']
  ];
  
  const ws = XLSX.utils.aoa_to_sheet(sampleData);
  
  // 컬럼 너비 설정
  ws['!cols'] = [
    { wch: 15 }, // 이름
    { wch: 10 }, // 성별
    { wch: 15 }  // 직군
  ];
  
  XLSX.utils.book_append_sheet(wb, ws, '학생정보');
  
  // 파일 다운로드
  XLSX.writeFile(wb, '학생정보_샘플.xlsx');
}

// 학생 정보 엑셀 업로드 및 파싱
function handleStudentExcelUpload(file) {
  const reader = new FileReader();
  
  reader.onload = function(e) {
    try {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: 'array' });
      
      // 첫 번째 시트 가져오기
      const firstSheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[firstSheetName];
      
      // JSON으로 변환
      const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
      
      // 헤더 제거 및 데이터 파싱
      const headers = jsonData[0];
      const studentDataArray = [];
      
      // 헤더 유효성 검사 (이름은 필수, 성별과 직군은 선택)
      if (!headers || headers.length < 1 || !headers[0].includes('이름')) {
        alert('엑셀 형식이 올바르지 않습니다.\n첫 번째 열은 "이름" 헤더여야 합니다.');
        return;
      }
      
      // 데이터 행 처리
      for (let i = 1; i < jsonData.length; i++) {
        const row = jsonData[i];
        if (row && row[0]) { // 이름이 있는 경우만 처리
          const name = String(row[0]).trim();
          const gender = String(row[1] || '').trim();
          const job = String(row[2] || '').trim();
          
          if (name) { // 이름만 있으면 처리
            let dataString = name;
            
            if (gender) {
              // 성별 정규화
              let normalizedGender = gender;
              if (['남', '남성', 'M', 'm', '남자'].includes(gender)) {
                normalizedGender = '남';
              } else if (['여', '여성', 'F', 'f', '여자'].includes(gender)) {
                normalizedGender = '여';
              } else {
                alert(`행 ${i+1}: 성별 형식이 올바르지 않습니다 (${gender})`);
                return;
              }
              
              dataString += `/${normalizedGender}`;
              
              if (job) {
                dataString += `/${job}`;
              }
            }
            
            studentDataArray.push(dataString);
          }
        }
      }
      
      if (studentDataArray.length === 0) {
        alert('유효한 학생 데이터가 없습니다.');
        return;
      }
      
      // textarea에 데이터 입력
      document.getElementById('student-data').value = studentDataArray.join('\n');
      
      // 업로드 상태 표시
      const statusDiv = document.getElementById('upload-status');
      statusDiv.textContent = `✅ ${studentDataArray.length}명의 학생 정보가 업로드되었습니다.`;
      statusDiv.style.color = '#48bb78';
      
    } catch (error) {
      console.error('엑셀 파싱 오류:', error);
      alert('엑셀 파일을 읽는 중 오류가 발생했습니다.');
      
      const statusDiv = document.getElementById('upload-status');
      statusDiv.textContent = '❌ 업로드 실패';
      statusDiv.style.color = '#f56565';
    }
  };
  
  reader.readAsArrayBuffer(file);
}

// 우선순위 드래그 앤 드롭 기능
function initializePriorityDragDrop() {
  const priorityList = document.getElementById('priority-list');
  
  if (priorityList && typeof Sortable !== 'undefined') {
    new Sortable(priorityList, {
      handle: '.drag-handle',
      animation: 150,
      ghostClass: 'priority-item-ghost',
      onStart: function(evt) {
        evt.item.classList.add('dragging');
      },
      onEnd: function(evt) {
        evt.item.classList.remove('dragging');
        updatePriorityOrder();
      }
    });
  } else {
    console.warn('Sortable not available or priority-list not found');
  }
}

// 우선순위 순서 업데이트
function updatePriorityOrder() {
  const items = document.querySelectorAll('.priority-item');
  priorityOrder = Array.from(items).map(item => item.dataset.type);
  console.log('우선순위 업데이트:', priorityOrder);
}

// 우선순위에 따른 가중치 계산
function getPriorityWeights() {
  const weights = {};
  const totalPriorities = priorityOrder.length;
  
  priorityOrder.forEach((type, index) => {
    // 첫 번째가 가장 중요 (가중치 높음), 마지막이 가장 덜 중요
    weights[type] = totalPriorities - index;
  });
  
  return weights;
}

// 체크박스 상태에 따른 우선순위 섹션 표시/숨김
function updatePrioritySectionVisibility() {
  const genderChecked = document.getElementById('balance-gender').checked;
  const jobChecked = document.getElementById('balance-job').checked;
  const prioritySection = document.getElementById('priority-section');
  
  if (genderChecked || jobChecked) {
    prioritySection.style.display = 'block';
    
    // 체크되지 않은 항목은 숨기기
    const genderItem = document.querySelector('[data-type="gender"]');
    const jobItem = document.querySelector('[data-type="job"]');
    
    if (genderItem) {
      genderItem.style.display = genderChecked ? 'flex' : 'none';
    }
    if (jobItem) {
      jobItem.style.display = jobChecked ? 'flex' : 'none';
    }
  } else {
    prioritySection.style.display = 'none';
  }
}

// DOM이 로드된 후 이벤트 리스너 등록
document.addEventListener('DOMContentLoaded', function() {
  // 제목 클릭시 페이지 새로고침
  document.getElementById('title').addEventListener('click', function() {
    location.reload();
  });
  
  // 홈으로 버튼 클릭시 공지 다시 표시
  document.getElementById('home-btn').addEventListener('click', function() {
    document.querySelector(".notice-text").style.display = "block";
  });
  
  // 단계별 다음 버튼
  document.getElementById('step1-next').addEventListener('click', function() {
    nextStep(1);
  });
  
  // 모드 토글 버튼
  document.getElementById('basic-mode-btn').addEventListener('click', function() {
    if (currentMode !== 'basic') {
      currentMode = 'basic';
      document.getElementById('basic-mode-btn').classList.add('active');
      document.getElementById('advanced-mode-btn').classList.remove('active');
      document.getElementById('basic-mode').style.display = 'block';
      document.getElementById('advanced-mode').style.display = 'none';
      document.getElementById('advanced-warning').style.display = 'none';
    }
  });
  
  document.getElementById('advanced-mode-btn').addEventListener('click', function() {
    if (currentMode !== 'advanced') {
      currentMode = 'advanced';
      document.getElementById('advanced-mode-btn').classList.add('active');
      document.getElementById('basic-mode-btn').classList.remove('active');
      document.getElementById('advanced-mode').style.display = 'block';
      document.getElementById('basic-mode').style.display = 'none';
      
      const warningEl = document.getElementById('advanced-warning');
      if (warningEl) {
        warningEl.style.display = 'block';
      }
      
      // 드래그 앤 드롭 초기화 (약간의 지연을 두고)
      setTimeout(() => {
        initializePriorityDragDrop();
        updatePrioritySectionVisibility();
      }, 100);
    }
  });
  
  document.getElementById('step2-next').addEventListener('click', function() {
    nextStep(2);
  });
  
  // 고급 모드 - 샘플 다운로드 버튼
  document.getElementById('download-sample').addEventListener('click', function() {
    downloadSampleExcel();
  });
  
  // 고급 모드 - 엑셀 업로드 버튼
  document.getElementById('upload-student-excel').addEventListener('click', function() {
    document.getElementById('studentExcelFile').click();
  });
  
  // 고급 모드 - 엑셀 파일 선택 시
  document.getElementById('studentExcelFile').addEventListener('change', function(e) {
    const file = e.target.files[0];
    if (file) {
      // 파일 크기 제한 (5MB)
      if (file.size > 5 * 1024 * 1024) {
        alert('파일 크기는 5MB를 초과할 수 없습니다.');
        e.target.value = '';
        return;
      }
      
      handleStudentExcelUpload(file);
    }
  });
  
  // 체크박스 변경 시 우선순위 섹션 업데이트
  document.getElementById('balance-gender').addEventListener('change', function() {
    updatePrioritySectionVisibility();
  });
  
  document.getElementById('balance-job').addEventListener('change', function() {
    updatePrioritySectionVisibility();
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