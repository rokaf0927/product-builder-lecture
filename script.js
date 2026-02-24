
document.addEventListener('DOMContentLoaded', () => {
    const generateBtn = document.getElementById('generate-btn');
    const resultContainer = document.getElementById('result');
    const pastDataContainer = document.getElementById('past-data-container');

    let lottoData = [];

    // Function to load and process the Excel file automatically
    async function loadAndProcessExcel() {
        try {
            const response = await fetch('로또 회차별 당첨번호.xlsx');
            if (!response.ok) {
                throw new Error(`네트워크 오류: ${response.status} ${response.statusText}. '로또 회차별 당첨번호.xlsx' 파일이 현재 폴더에 있는지 확인해주세요.`);
            }
            const arrayBuffer = await response.arrayBuffer();
            const data = new Uint8Array(arrayBuffer);
            const workbook = XLSX.read(data, { type: 'array' });
            const firstSheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[firstSheetName];
            const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1, raw: false });

            if (jsonData.length < 2) {
                throw new Error("Excel 파일에 데이터가 부족합니다.");
            }

            // Find the starting row of the actual data
            let dataStartIndex = -1;
            for (let i = 0; i < jsonData.length; i++) {
                if (jsonData[i] && jsonData[i].length > 1 && jsonData[i][1] && !isNaN(parseInt(String(jsonData[i][1]).replace(/,/g, ''), 10))) {
                    dataStartIndex = i;
                    break;
                }
            }

            if (dataStartIndex === -1) {
                throw new Error("처리할 수 있는 데이터가 없습니다. Excel 파일의 B열에 숫자 형식의 회차 데이터가 있는지 확인해주세요.");
            }

            lottoData = parseExcelData(jsonData.slice(dataStartIndex));

            if (lottoData.length === 0) {
                throw new Error("처리할 수 있는 데이터가 없습니다. Excel 파일의 데이터 형식을 확인해주세요.");
            }

            displayPastData(lottoData);
            generateBtn.disabled = false; // Enable button after data is loaded
            resultContainer.innerHTML = ""; 

        } catch (error) {
            console.error("파일 처리 중 오류:", error);
            pastDataContainer.innerHTML = `<p style="color: red;">${error.message}</p>`;
            generateBtn.disabled = true;
        }
    }

    function parseExcelData(dataRows) {
        // B열: 회차 (인덱스 1)
        const roundIndex = 1;
        // C열~H열: 당첨 번호 (인덱스 2~7)
        const winNumIndices = [2, 3, 4, 5, 6, 7];
        // I열: 보너스 번호 (인덱스 8)
        const bonusIndex = 8;

        const parsedData = dataRows.map((row, rowIndex) => {
            try {
                if (row.length <= bonusIndex) {
                    console.warn(`경고: 데이터 행 ${rowIndex + 1}에 데이터가 충분하지 않습니다. 건너뜁니다.`);
                    return null;
                }

                const round = parseInt(String(row[roundIndex]).replace(/,/g, ''), 10);
                const win_nums = winNumIndices.map(i => parseInt(String(row[i]).replace(/,/g, ''), 10));
                const bonus_num = parseInt(String(row[bonusIndex]).replace(/,/g, ''), 10);

                if (isNaN(round) || win_nums.some(isNaN) || isNaN(bonus_num)) {
                     console.warn(`경고: 데이터 행 ${rowIndex + 1}에 유효하지 않은 숫자 데이터가 있습니다. 건너뜁니다.`);
                     return null;
                }

                return {
                    round: round,
                    win_nums: win_nums.sort((a, b) => a - b),
                    bonus_num: bonus_num
                };
            } catch (e) {
                console.error(`데이터 행 ${rowIndex + 1} 처리 중 오류 발생:`, e);
                return null;
            }
        }).filter(Boolean); // Filter out any null (invalid) entries

        // Sort by round number in descending order
        return parsedData.sort((a, b) => b.round - a.round);
    }

    // Event listener for the generate button
    generateBtn.addEventListener('click', () => {
        if (lottoData.length === 0) {
            resultContainer.innerHTML = `<p style="color: red;">데이터를 먼저 불러와야 합니다.</p>`;
            return;
        }

        // 각 자리별(1~6번째) 숫자 출현 빈도를 저장할 배열
        const positionCounts = Array(6).fill(null).map(() => new Map());

        // 과거 데이터를 분석하여 각 자리별 숫자 출현 빈도 계산
        // win_nums는 오름차순으로 정렬되어 있음
        lottoData.forEach(round => {
            round.win_nums.forEach((num, index) => {
                if (index < 6) { // 6개의 당첨 번호에 대해서만 처리
                    const counts = positionCounts[index];
                    counts.set(num, (counts.get(num) || 0) + 1);
                }
            });
        });

        const recommendedNumbers = [];
        const usedNumbers = new Set();

        // 각 자리에서 가장 많이 나온 숫자를 선택 (중복 제외)
        for (let i = 0; i < 6; i++) {
            const sortedByCount = [...positionCounts[i].entries()].sort((a, b) => b[1] - a[1]);
            
            let foundNumber = false;
            // 아직 선택되지 않은 숫자 중에서 가장 빈도가 높은 것을 찾음
            for (const [num, count] of sortedByCount) {
                if (!usedNumbers.has(num)) {
                    recommendedNumbers.push(num);
                    usedNumbers.add(num);
                    foundNumber = true;
                    break;
                }
            }

            // 만약 해당 자리의 모든 상위 빈도 숫자가 이미 선택되었다면,
            // 전체 1-45 숫자 중에서 아직 선택되지 않은 숫자로 채움 (fallback)
            if (!foundNumber) {
                for (let num = 1; num <= 45; num++) {
                    if (!usedNumbers.has(num)) {
                        recommendedNumbers.push(num);
                        usedNumbers.add(num);
                        break;
                    }
                }
            }
        }

        // 최종 추천 번호를 오름차순으로 정렬하여 표시
        displayRecommendedNumbers(recommendedNumbers.sort((a, b) => a - b));
    });

    function displayPastData(data) {
        // Create a copy of the data array to avoid modifying the original
        const dataCopy = [...data];

        // Fisher-Yates shuffle algorithm to shuffle the data array
        for (let i = dataCopy.length - 1; i > 0; i--) {
            const j = Math.floor(Math.random() * (i + 1));
            [dataCopy[i], dataCopy[j]] = [dataCopy[j], dataCopy[i]];
        }

        const randomData = dataCopy.slice(0, 10);

        const table = document.createElement('table');
        table.innerHTML = `
            <thead>
                <tr>
                    <th>회차</th>
                    <th>당첨 번호</th>
                    <th>보너스</th>
                </tr>
            </thead>
            <tbody>
                ${randomData.map(round => `
                    <tr>
                        <td>${round.round}</td>
                        <td>${round.win_nums.join(', ')}</td>
                        <td>${round.bonus_num}</td>
                    </tr>
                `).join('')}
            </tbody>
        `;
        pastDataContainer.innerHTML = ''; 
        pastDataContainer.appendChild(table);
    }

    function displayRecommendedNumbers(numbers) {
        resultContainer.innerHTML = '';
        numbers.forEach(num => {
            const ball = document.createElement('div');
            ball.className = 'lotto-ball';
            ball.textContent = num;
            resultContainer.appendChild(ball);
        });
    }

    // Automatically load the data when the page loads
    loadAndProcessExcel();
});
