
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
            pastDataContainer.innerHTML = `<p style=\"color: red;\">${error.message}</p>`;
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
            resultContainer.innerHTML = `<p style=\"color: red;\">데이터를 먼저 불러와야 합니다.</p>`;
            return;
        }

        const numberCounts = new Map();
        for (let i = 1; i <= 45; i++) {
            numberCounts.set(i, 0);
        }

        lottoData.forEach(round => {
            round.win_nums.forEach(num => {
                if (numberCounts.has(num)) {
                    numberCounts.set(num, numberCounts.get(num) + 1);
                }
            });
        });

        const sortedNumbers = [...numberCounts.entries()].sort((a, b) => b[1] - a[1]);
        const top6Numbers = sortedNumbers.slice(0, 6).map(entry => entry[0]);

        displayRecommendedNumbers(top6Numbers.sort((a, b) => a - b));
    });

    function displayPastData(data) {
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
                ${data.map(round => `
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
