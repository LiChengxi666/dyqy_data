let data = []; // 存储所有工作表的观测数据
let allSites = []; // 存储所有观测点的名称

// 自动读取数据
const dataPath = 'data.xlsx'; // 在这里指定服务器上数据文件的路径

// 使用 fetch 自动读取数据
fetch(dataPath)
    .then(response => response.arrayBuffer())
    .then(arrayBuffer => {
        const workbook = XLSX.read(arrayBuffer, { type: 'array' });
        processData(workbook);
        searchData(); // 页面加载时自动展示全部数据
    })
    .catch(error => {
        console.error("数据加载失败:", error);
    });

// 处理数据
function processData(workbook) {
    data = [];
    allSites = [];
    workbook.SheetNames.forEach(sheetName => {
        const worksheet = workbook.Sheets[sheetName];
        const json = XLSX.utils.sheet_to_json(worksheet, { header: 1 }); // 获取工作表的数据
        
        if (json.length > 1) {
            const siteData = {
                siteName: sheetName,  // 工作表名称即为观测点名称
                birds: json.slice(1).map(row => {  // 跳过第一行表头
                    return {
                        birdID: row[0],
                        chineseName: row[1],
                        latinName: row[2],
                        englishName: row[3],
                        order: row[4],
                        family: row[5],
                        count: row[6],
                        classification: `${row[4]} - ${row[5]}` // 拼接目和科作为分类
                    };
                })
            };
            data.push(siteData);
            allSites.push(sheetName);  // 记录观测点名称
        }
    });

    // 更新观测点下拉框
    const siteSelect = document.getElementById('site');
    siteSelect.innerHTML = '';  // 清空下拉框
    const allOption = document.createElement('option');
    allOption.value = "";
    allOption.textContent = "全部观测点";  // 添加"全部观测点"选项
    siteSelect.appendChild(allOption);
    allSites.forEach(site => {
        const option = document.createElement('option');
        option.value = site;
        option.textContent = site;
        siteSelect.appendChild(option);
    });
}

// 搜索数据
function searchData() {
    const searchKeyword = document.getElementById("searchKeyword").value; // 获取选择的关键词
    const searchText = document.getElementById("searchText").value.toLowerCase(); // 获取用户输入的检索文本
    const selectedSite = document.getElementById("site").value.toLowerCase(); // 获取选择的观测点

    let resultHTML = '';

    // 遍历所有工作表（观测点）
    data.forEach(siteData => {
        const siteName = siteData.siteName.toLowerCase();

        // 如果选择了观测点且不匹配，或者选择“全部观测点”，都进行处理
        if (selectedSite && selectedSite !== "" && siteName.indexOf(selectedSite) === -1) {
            return; // 如果指定了观测点，且不匹配，则跳过
        }

        // 遍历该观测点下的所有鸟种数据
        siteData.birds.forEach(bird => {
            const searchValue = bird[searchKeyword].toLowerCase(); // 按选择的关键词检索

            // 如果鸟种信息不符合搜索条件，且检索文本不为空
            if (searchText && searchValue.indexOf(searchText) === -1) {
                return;
            }

            resultHTML += `
                <tr>
                    <td>${siteName}</td>
                    <td>${bird.birdID}</td>
                    <td>${bird.chineseName}</td>
                    <td class="latinName">${bird.latinName}</td>
                    <td>${bird.englishName}</td>
                    <td>${bird.classification}</td>
                    <td>${bird.count}</td>
                </tr>
            `;
        });
    });

    // 如果没有找到结果
    if (resultHTML === '') {
        resultHTML = '<tr><td colspan="7">未找到匹配的结果。</td></tr>';
    }

    // 显示结果
    document.querySelector("#resultTable tbody").innerHTML = resultHTML;
}

// 导出为Excel
function exportToExcel() {
    const ws = XLSX.utils.table_to_sheet(document.getElementById('resultTable'));
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, '检索结果');
    XLSX.writeFile(wb, 'bird_observation_data.xlsx');
}