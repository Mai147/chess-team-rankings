const RANK = ["hạng", "rk."];
const CLUB_NAME = ["fed", "lđ"];
const POINTS = ["pts.", "điểm"];
const NAME = ["name", "tên"];
const RANK_MODE = "rank";
const POINT_MODE = "point";

const submitBtn = document.querySelector(".submit-btn");
const submitBtnContent = document.querySelector(".submit-btn-content");
const resultContainer = document.querySelector(".result-container");
const playerNumberInput = document.querySelector(".player-number-input");
const modeOption = document.querySelector(".mode-option");
const optionTypeBar = document.querySelector(".option-type-bar");
const optionContent = document.querySelector(".option-content");

const getSheetRows = (sheet) => {
    let rows = [];
    const range = XLSX.utils.decode_range(sheet["!ref"]);
    for (let rowNum = range.s.r; rowNum <= range.e.r; rowNum++) {
        const row = [];
        // Iterate through columns
        for (let colNum = range.s.c; colNum <= range.e.c; colNum++) {
            const cell = getCellContent(sheet, rowNum, colNum);
            row.push(cell ? cell.v : undefined); // Push cell value or undefined
        }
        rows = [...rows, row];
    }
    return rows;
};

const getCellContent = (sheet, rowNum, colNum) => {
    const cellAddress = { r: rowNum, c: colNum };
    const cellRef = XLSX.utils.encode_cell(cellAddress);
    const cell = sheet[cellRef];
    return cell;
};

const getTitleRowIdx = (rows) => {
    const titleRow = rows.find((row) =>
        RANK.includes(row[0]?.toLowerCase()?.trim())
    );
    return rows.indexOf(titleRow);
};

const getClubNameColIdx = (titleRow) => {
    const item = titleRow.filter((value) =>
        CLUB_NAME.includes(value?.toLowerCase()?.trim())
    )[0];
    return titleRow.indexOf(item);
};

const getPointsColIdx = (titleRow) => {
    const item = titleRow.filter((value) =>
        POINTS.includes(value?.toLowerCase()?.trim())
    )[0];
    return titleRow.indexOf(item);
};

const getNameColIdx = (titleRow) => {
    const item = titleRow.filter((value) =>
        NAME.includes(value?.toLowerCase()?.trim())
    )[0];
    return titleRow.indexOf(item);
};

const groupByClubName = (
    rows,
    clubNameIdx,
    pointsIdx,
    nameIdx,
    numberOfPlayers
) => {
    let result = [];
    rows.forEach((row) => {
        const name = row[clubNameIdx];
        const item = result.find((res) => res.name === name);
        if (!item) {
            result = [
                ...result,
                initGroupByRow(name, rows, row, pointsIdx, nameIdx),
            ];
        } else {
            const idx = result.indexOf(item);
            if (result[idx].players.length < numberOfPlayers) {
                result[idx] = getNewRankPlayerRes(
                    result[idx],
                    rows,
                    row,
                    pointsIdx,
                    nameIdx
                );
            }
        }
    });
    return result;
};

const initGroupByRow = (name, rows, row, pointsIdx, nameIdx) => {
    const point = formatPoint(row[pointsIdx]);
    const rank = findSameRank(rows, row);
    const player = {
        rank,
        point: row[pointsIdx],
        name: row[nameIdx],
    };
    return {
        name,
        players: [player],
        totalRank: rank,
        totalPoints: point,
    };
};

const getNewRankPlayerRes = (res, rows, row, pointsIdx, nameIdx) => {
    const point = formatPoint(row[pointsIdx]);
    const rank = findSameRank(rows, row);
    const player = {
        rank,
        point: row[pointsIdx],
        name: row[nameIdx],
    };
    return {
        ...res,
        players: [...(res.players ?? []), player],
        totalRank: res.totalRank + rank,
        totalPoints: res.totalPoints + point,
    };
};

const findSameRank = (rows, row) => {
    if (row[0]) return Number(row[0]);
    const idx = rows.indexOf(row);
    for (let i = idx; i >= 0; i--) {
        if (rows[i][0]) return Number(rows[i][0]);
    }
};

const formatPoint = (point) => {
    if (!isNaN(Number(point))) return Number(point);
    return Number(point.slice(0, -1));
};

const calculateResult = (
    rows,
    clubNameIdx,
    pointsIdx,
    nameIdx,
    numberOfPlayers,
    mode
) => {
    const groupByResult = groupByClubName(
        rows,
        clubNameIdx,
        pointsIdx,
        nameIdx,
        numberOfPlayers
    );
    const result = groupByResult.filter(
        (res) => res.players.length >= numberOfPlayers
    );
    if (mode === RANK_MODE) result.sort(compareRankFunction);
    else if (mode === POINT_MODE) result.sort(comparePointFunction);
    return result;
};

const compareRankFunction = (a, b) => {
    if (a.totalRank < b.totalRank) return -1;
    if (a.totalRank > b.totalRank) return 1;
    if (a.totalPoints > b.totalPoints) return -1;
    if (a.totalPoints < b.totalPoints) return 1;
    if (Number(a.players[0][0]) < Number(b.players[0][0])) return -1;
    if (Number(a.players[0][0]) > Number(b.players[0][0])) return 1;
    return 0;
};

const comparePointFunction = (a, b) => {
    if (a.totalPoints > b.totalPoints) return -1;
    if (a.totalPoints < b.totalPoints) return 1;
    if (a.totalRank < b.totalRank) return -1;
    if (a.totalRank > b.totalRank) return 1;
    if (Number(a.players[0][0]) < Number(b.players[0][0])) return -1;
    if (Number(a.players[0][0]) > Number(b.players[0][0])) return 1;
    return 0;
};

const getPlayerResultRows = (rows, titleRowIdx) => {
    const rankRows = rows.slice(titleRowIdx + 1);
    let endIdx = 0;
    for (let i = 0; i < rankRows.length; i++) {
        if (!rankRows[i].some((item) => item)) {
            endIdx = i;
            break;
        }
    }
    return rankRows.slice(0, endIdx);
};

const calculateTeamResult = (sheet, numberOfPlayers, mode) => {
    const rows = getSheetRows(sheet);
    const titleRowIdx = getTitleRowIdx(rows);
    const titleRow = rows[titleRowIdx];
    const clubNameIdx = getClubNameColIdx(titleRow);
    const pointsIdx = getPointsColIdx(titleRow);
    const nameIdx = getNameColIdx(titleRow);
    const playerResultRows = getPlayerResultRows(rows, titleRowIdx);
    const res = calculateResult(
        playerResultRows,
        clubNameIdx,
        pointsIdx,
        nameIdx,
        numberOfPlayers,
        mode
    );
    return res;
};

const renderResult = (res) => {
    resultContainer.innerHTML = `
        <h3>BẢNG XẾP HẠNG ĐỒNG ĐỘI</h3>
            <div class="result-table scroll">
                <div class="result-item result-header">
                    <div>
                        <span>Hạng</span>
                    </div>
                    <div>
                        <span>Đội</span>
                    </div>
                    <div>
                        <span>Vận động viên</span>
                    </div>
                    <div>
                        <span>Hạng cá nhân</span>
                    </div>
                    <div>
                        <span>Tổng hạng</span>
                    </div>
                    <div>
                        <span>Điểm cá nhân</span>
                    </div>
                    <div>
                        <span>Tổng điểm</span>
                    </div>
                </div>
                ${res
                    .map(
                        (item, idx) => `
                    <div class="result-item">
                      <span>${idx + 1}</span>
                      <span>${item.name}</span>
                      <div class="result-column">
                      ${item.players
                          .map((player) => `<span>${player.name}</span>`)
                          .join("")}
                      </div>
                      <div class="result-column">
                      ${item.players
                          .map((player) => `<span>${player.rank}</span>`)
                          .join("")}
                      </div>
                      <span>${item.totalRank}</span>
                      <div class="result-column">
                      ${item.players
                          .map((player) => `<span>${player.point}</span>`)
                          .join("")}
                      </div>
                      <span>${item.totalPoints}</span>
                  </div>
                  `
                    )
                    .join("")}
            </div>`;
};

const formatChessResultLink = (value) => {
    const roundIdx = value.indexOf("rd=");
    let url = value;
    const firstParamIdx = value.indexOf("?") + 1;
    url = url.slice(0, firstParamIdx);
    url += "lan=1&art=1&zeilen=0&prt=4&excel=2010&";
    if (!roundIdx) {
        url += "rd=9";
    } else {
        value.slice(roundIdx);
    }
    return url;
};

const calculateResultFromFile = async (file, numberOfPlayers, mode) => {
    const reader = new FileReader();
    reader.onload = (e) => {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, {
            type: "array",
        });

        const sheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[sheetName];

        // Iterate through rows
        const res = calculateTeamResult(sheet, numberOfPlayers, mode);
        renderResult(res);
    };
    reader.onerror = function (ex) {
        alert("Có lỗi xảy ra. Vui lòng thử lại");
        console.log(ex);
    };

    reader.readAsArrayBuffer(file);
};

const calculateResultFromLink = async (value, numberOfPlayers, mode) => {
    const proxyUrl = `https://api.allorigins.win/get?url=`;
    const url = formatChessResultLink(value);
    const res = await fetch(proxyUrl + encodeURIComponent(url));
    const data = await res.json();
    const base64Content = data.contents.replace(
        "data:application/vnd.ms-excel;base64,",
        ""
    );
    const workbook = XLSX.read(
        base64Content.replace(/_/g, "/").replace(/-/g, "+"),
        { type: "base64" }
    );

    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];

    // Iterate through rows
    const teamRes = calculateTeamResult(sheet, numberOfPlayers, mode);
    renderResult(teamRes);
};

// Events
const uploadFile = () => {
    const fileInput = document.getElementById("file-input");
    fileInput.click();
};

const submit = async () => {
    const optionTypeBarActiveItem = optionTypeBar.querySelector(
        ".option-type-bar-item.active"
    );
    const currentType = optionTypeBarActiveItem.classList.contains("file-type")
        ? "file"
        : "link";
    const loader = document.createElement("div");
    loader.classList.add("loader");
    submitBtnContent.appendChild(loader);
    submitBtn.disabled = true;
    if (currentType === "file") {
        const fileInput = document.getElementById("file-input");
        const file = fileInput.files[0];
        const numberOfPlayers = playerNumberInput.value;
        const mode = modeOption.value;
        if (file) {
            await calculateResultFromFile(file, numberOfPlayers, mode);
        }
    } else {
        const linkInput = document.querySelector(".link-input input");
        const value = linkInput.value;
        const numberOfPlayers = playerNumberInput.value;
        const mode = modeOption.value;
        if (value) {
            await calculateResultFromLink(value, numberOfPlayers, mode);
        }
    }
    submitBtnContent.removeChild(loader);
    submitBtn.disabled = false;
};

const handleFile = () => {
    const fileInput = document.getElementById("file-input");
    const file = fileInput.files[0];
    const numberOfPlayers = playerNumberInput.value;
    const mode = modeOption.value;
    if (file) calculateResultFromFile(file, numberOfPlayers, mode);
};

const onInputFileChange = (e) => {
    const fileInput = document.getElementById("file-input");
    const fileName = document.querySelector(".file-name");
    const file = fileInput.files[0];
    if (file) {
        const name = file.name;
        fileName.innerHTML = name;
        submitBtn.classList.add("active");
        resultContainer.innerHTML = "";
    } else {
        submitBtn.classList.remove("active");
    }
};

const onPlayerNumberInputChange = (e) => {
    const value = e.target.value;
    if (!isValidNumberPlayer(value)) {
        playerNumberInput.value = value.slice(0, -1);
    }
};

const onLinkInputChange = (e) => {
    const value = e.target.value;
    if (value) {
        submitBtn.classList.add("active");
        resultContainer.innerHTML = "";
    } else {
        submitBtn.classList.remove("active");
    }
};

const onTypeBarItemChange = () => {
    const optionTypeBarActiveItem = optionTypeBar.querySelector(
        ".option-type-bar-item.active"
    );
    const optionTypeBarUnActiveItem = optionTypeBar.querySelector(
        ".option-type-bar-item:not(.active)"
    );
    const switchType = optionTypeBarUnActiveItem.classList.contains("file-type")
        ? "file"
        : "link";
    optionTypeBarActiveItem.classList.remove("active");
    optionTypeBarUnActiveItem.classList.add("active");
    if (switchType === "file") {
        optionContent.innerHTML = `<input type="file" id="file-input" hidden accept="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" onchange="onInputFileChange()">
      <div class="file-group">
          <button class="file-btn" onclick="uploadFile()">
              <span class="file-btn-text">File xếp hạng</span>
              <i class="fa-solid fa-arrow-up-from-bracket"></i>
          </button>
          <span class="file-name"></span>
      </div>`;
    } else {
        optionContent.innerHTML = `<div class="form-group link-input">
          <input style="width: 250px;" type="text" placeholder="Link Chess Results (ván 9)" oninput="onLinkInputChange(event)">
      </div>`;
    }
};

// Utilities
const isValidNumberPlayer = (value) => {
    if (isNaN(Number(value))) return false;
    const numberValue = Number(value);
    if (!Number.isInteger(numberValue)) return false;
    if (numberValue <= 1) return false;
    return true;
};
