let rows = 100;
let cols = 26;

let addressColContainer = document.querySelector(".address-col-container");
let addressRowContainer = document.querySelector(".address-row-container");
let cellsContainer = document.querySelector(".cells-container");
let addressBar = document.querySelector(".address-bar");


for(let i = 0; i<rows; i++){
    let addressCol = document.createElement("div");
    addressCol.setAttribute("class", "address-col");
    addressCol.innerText = i+1;
    addressColContainer.appendChild(addressCol);
}
for(let i = 0; i<cols; i++){
    let addressRow = document.createElement("div");
    addressRow.setAttribute("class", "address-row");
    addressRow.innerText = String.fromCharCode(65 + i);
    addressRowContainer.appendChild(addressRow);
}

for(let row = 0; row < rows; row++){
    let cellRow = document.createElement("div");
    cellRow.setAttribute("class", "cell-row");
    for(let col = 0; col < cols; col++){
        let cell = document.createElement("div");
        cell.setAttribute("class", "cell");
        cell.setAttribute("contenteditable", "true");
        cell.setAttribute("spellcheck", "false");

        cell.setAttribute("rowId", row);
        cell.setAttribute("colId", col);

        cellRow.appendChild(cell);
        addressBarDisplay(cell, row, col);
    }
    cellsContainer.appendChild(cellRow);
}

function addressBarDisplay(cell, row, col){
    cell.addEventListener('click', (e) => {
        let rowId = row+1;
        let colId = String.fromCharCode(65+col);
        addressBar.value = `${colId}${rowId}`;
    })
}



let allGraphComponentMatrix = [];
let graphComponentMatrix = [];

function isGraphCycle(){
    let visited = [], dfsVisited = [];

    for(let i = 0; i<rows; i++){
        let visitedRow = [], dfsVisitedRow = [];
        for(let j = 0; j<cols; j++){
            visitedRow.push(false);
            dfsVisitedRow.push(false);
        }
        visited.push(visitedRow);
        dfsVisited.push(dfsVisitedRow);
    }
    for(let i = 0; i<rows; i++){
        for(let j = 0; j<cols; j++){
            if(!visited[i][j]){
                let response = dfsCycleDetection(i, j, visited, dfsVisited);
                if(response) return [i, j];
            }
        }
    }
    return null;
}

function dfsCycleDetection(row, col, visited, dfsVisited){
    visited[row][col] = true;
    dfsVisited[row][col] = true;

    for(let children = 0; children < graphComponentMatrix[row][col].length; children++){
        let [rowId, colId] = graphComponentMatrix[row][col][children];
        if(!visited[rowId][colId]){
            let response = dfsCycleDetection(rowId, colId, visited, dfsVisited);
            if(response) return true; // found cycle
        }
        else if(dfsVisited[rowId][colId]) return true;
    }

    dfsVisited[row][col] = false;
    return false;
}



let sheetFolderContainer = document.querySelector(".sheet-folder-container");
let sheetAddBtn = document.querySelector(".sheet-add-icon");

sheetAddBtn.addEventListener("click", (e) => {
    let sheet = document.createElement("div");
    sheet.setAttribute("class", "sheet-folder");

    let allSheetFolders = document.querySelectorAll(".sheet-folder");
    sheet.setAttribute("id", allSheetFolders.length);

    sheet.innerHTML = `
        <div class="sheet-content">Sheet ${allSheetFolders.length + 1}</div> 
    `

    sheetFolderContainer.appendChild(sheet);
    sheet.scrollIntoView();

    createSheetDB();
    createGraphComponentMatrix();
    handleActiveSheet(sheet);
    handleSheetRemoval(sheet);
    sheet.click();
})

function handleSheetRemoval(sheet){
    sheet.addEventListener("mousedown", (e) => {
        // checking for right click only
        if(e.button !== 2) return;

        let allSheetFolders = document.querySelectorAll(".sheet-folder");
        if(allSheetFolders.length === 1){
            alert("Cannot delete, you need atleast one sheet.");
            return;
        }

        let response = confirm("Do you want delete sheet permanently ?");
        if(response === false) return;

        let sheetId = Number(sheet.getAttribute("id"));
        allSheetsDB.splice(sheetId, 1);
        allGraphComponentMatrix.splice(sheetId,  1);

        // make previous sheet active
        let activeIndex = Math.max(0, Number(sheetId - 1));
        handleSheetUIRemoval(sheet, activeIndex);
        sheetDB = allSheetsDB[activeIndex];
        graphComponentMatrix = allGraphComponentMatrix[activeIndex];
        handleSheetProperties();
    })
}

function handleSheetUIRemoval(sheet, activeIndex){
    sheet.remove();
    let allSheetFolders = document.querySelectorAll(".sheet-folder");
    for(let i = 0; i<allSheetFolders.length; i++){
        allSheetFolders[i].setAttribute("id", i);
        let sheetContent = allSheetFolders[i].querySelector(".sheet-content");
        sheetContent.innerHTML = `Sheet ${i+1}`;
        allSheetFolders[i].style.backgroundColor = "transparent";
    }

    allSheetFolders[activeIndex].style.backgroundColor = "#ced6e0";
}
function handleSheetUI(sheet){
    let allSheetFolder = document.querySelectorAll(".sheet-folder");
    for(let i = 0; i<allSheetFolder.length; i++){
        allSheetFolder[i].style.backgroundColor = "transparent";
    }
    sheet.style.backgroundColor = "#ced6e0";
}
function handleActiveSheet(sheet){
    sheet.addEventListener("click", (e) => {
        let sheetId = Number(sheet.getAttribute("id"));
        handleSheetDB(sheetId);
        handleSheetProperties(sheetId);
        handleSheetUI(sheet);
    })
}

function handleSheetDB(sheetId){
    sheetDB = allSheetsDB[sheetId];
    graphComponentMatrix = allGraphComponentMatrix[sheetId];
}

function handleSheetProperties(sheetId){
    for(let i = 0; i<rows; i++){
        for(let j = 0; j<cols; j++){
            let cell = document.querySelector(`.cell[rowId="${i}"][colId="${j}"]`);
            cell.click();
        }
    }
    // By default, first cell should be active
    let firstCell = document.querySelector(".cell");
    firstCell.click();
}

function defaultCellProps(){
    return {
        bold: false,
        italic: false,
        underline: false,
        alignment: "left",
        fontFamily: "monospace",
        fontSize: "14",
        fontColor: "#000000",
        BGColor: "#ecf0f1",
        value: "",
        formula: "",
        children: []
    };
}
function createSheetDB(){
    let sheetDB = [];

    for(let i = 0; i<rows; i++){
        let sheetRow = [];
        for(let j = 0; j<cols; j++){
            let cellProp = defaultCellProps();
            sheetRow.push(cellProp)
        }
        sheetDB.push(sheetRow);
    }
    allSheetsDB.push(sheetDB);
}

function createGraphComponentMatrix(){
    let graphComponentMatrix = [];

    for(let i = 0; i<rows; i++){
        let row = [];
        for(let j = 0; j<cols; j++){
            row.push([]);
        }
        graphComponentMatrix.push(row);
    }

    allGraphComponentMatrix.push(graphComponentMatrix);
}





// Storage
let allSheetsDB = [];
let sheetDB = [];

{
    let sheetAddBtn = document.querySelector(".sheet-add-icon");
    sheetAddBtn.click();
}

// Selectors for cell properties
let bold = document.querySelector(".bold");
let italic = document.querySelector(".italic");
let underline = document.querySelector(".underline");
let fontSize = document.querySelector(".font-size-prop");
let fontFamily = document.querySelector(".font-family-prop");
let fontColor = document.querySelector(".font-color-prop");
let BGColor = document.querySelector(".bg-color-prop");
let alignment = document.querySelectorAll(".alignment");
let leftAlign = alignment[0];
let centerAlign = alignment[1];
let rightAlign = alignment[2];
let activeColorProp = "#d1d8e0";
let inactiveColorProp =  "#ecf0f1";

// Attach Property listeners
bold.addEventListener("click", (e) => {
    let address = addressBar.value;
    let [cell, cellProp] = getActiveCellAndProps(address);

    //Modification
    cellProp.bold = !cellProp.bold;
    cell.style.fontWeight = cellProp.bold ? "bold" : "normal";
    bold.style.backgroundColor = cellProp.bold ? activeColorProp : inactiveColorProp;
})
italic.addEventListener("click", (e) => {
    let address = addressBar.value;
    let [cell, cellProp] = getActiveCellAndProps(address);

    //Modification
    cellProp.italic = !cellProp.italic;
    cell.style.fontStyle = cellProp.italic ? "italic" : "normal";
    italic.style.backgroundColor = cellProp.italic ? activeColorProp : inactiveColorProp
})
underline.addEventListener("click", (e) => {
    let address = addressBar.value;
    let [cell, cellProp] = getActiveCellAndProps(address);

    //Modification
    cellProp.underline = !cellProp.underline;
    cell.style.textDecoration = cellProp.underline ? "underline" : "unset";
    underline.style.backgroundColor = cellProp.underline ? activeColorProp : inactiveColorProp
})
fontSize.addEventListener("change", (e) => {
    let address = addressBar.value;
    let [cell, cellProp] = getActiveCellAndProps(address);

    //Modification
    cellProp.fontSize = fontSize.value;
    cell.style.fontSize = cellProp.fontSize + "px";
    fontSize.value = cellProp.fontSize;
})
fontFamily.addEventListener("change", (e) => {
    let address = addressBar.value;
    let [cell, cellProp] = getActiveCellAndProps(address);

    //Modification
    cellProp.fontFamily = fontFamily.value;
    cell.style.fontFamily = cellProp.fontFamily;
    fontFamily.value = cellProp.fontFamily;
})
fontColor.addEventListener("change", (e) => {
    let address = addressBar.value;
    let [cell, cellProp] = getActiveCellAndProps(address);

    //Modification
    cellProp.fontColor = fontColor.value;
    cell.style.color = cellProp.fontColor;
    fontColor.value = cellProp.fontColor;
})
BGColor.addEventListener("change", (e) => {
    let address = addressBar.value;
    let [cell, cellProp] = getActiveCellAndProps(address);

    //Modification
    cellProp.BGColor = BGColor.value;
    cell.style.backgroundColor = cellProp.BGColor;
    BGColor.value = cellProp.BGColor;
})
alignment.forEach((alignElem) => {
    alignElem.addEventListener('click', (e) => {
        let address = addressBar.value;
        let [cell, cellProp] = getActiveCellAndProps(address);

        let alignValue = e.target.classList[0];
        cellProp.alignment = alignValue;
        cell.style.textAlign = cellProp.alignment;
        switch (alignValue){
            case "left":
                leftAlign.style.backgroundColor = activeColorProp;
                centerAlign.style.backgroundColor = inactiveColorProp;
                rightAlign.style.backgroundColor = inactiveColorProp;
                break;
            case "center":
                leftAlign.style.backgroundColor = inactiveColorProp;
                centerAlign.style.backgroundColor = activeColorProp;
                rightAlign.style.backgroundColor = inactiveColorProp;
                break;
            case "right":
                leftAlign.style.backgroundColor = inactiveColorProp;
                centerAlign.style.backgroundColor = inactiveColorProp;
                rightAlign.style.backgroundColor = activeColorProp;
                break;
        }
    })
})

let allCells = document.querySelectorAll(".cell");
for(let i = 0; i<allCells.length; i++){
    defaultCellProperties(allCells[i]);
}

function defaultCellProperties(cell){
    cell.addEventListener('click', (e) => {
        let address = addressBar.value;
        let [rowId, colId] = decodeRowIdColId(address);
        let cellProp = sheetDB[rowId][colId];
        updateCellPropsUI(cell, cellProp);
    })
}

function updateCellPropsUI(cell, cellProp){
    cell.style.fontWeight = cellProp.bold ? "bold" : "normal";
    cell.style.fontStyle = cellProp.italic ? "italic" : "normal";
    cell.style.textDecoration = cellProp.underline ? "underline" : "unset";
    cell.style.fontSize = cellProp.fontSize + "px";
    cell.style.fontFamily = cellProp.fontFamily;
    cell.style.color = cellProp.fontColor;
    cell.style.backgroundColor = cellProp.BGColor;
    cell.style.textAlign = cellProp.alignment;

    bold.style.backgroundColor = cellProp.bold ? activeColorProp : inactiveColorProp;
    italic.style.backgroundColor = cellProp.italic ? activeColorProp : inactiveColorProp
    underline.style.backgroundColor = cellProp.underline ? activeColorProp : inactiveColorProp
    fontSize.value = cellProp.fontSize;
    fontColor.value = cellProp.fontColor;
    fontFamily.value = cellProp.fontFamily;
    BGColor.value = cellProp.BGColor;
    switch (cellProp.alignment){
        case "left":
            leftAlign.style.backgroundColor = activeColorProp;
            centerAlign.style.backgroundColor = inactiveColorProp;
            rightAlign.style.backgroundColor = inactiveColorProp;
            break;
        case "center":
            leftAlign.style.backgroundColor = inactiveColorProp;
            centerAlign.style.backgroundColor = activeColorProp;
            rightAlign.style.backgroundColor = inactiveColorProp;
            break;
        case "right":
            leftAlign.style.backgroundColor = inactiveColorProp;
            centerAlign.style.backgroundColor = inactiveColorProp;
            rightAlign.style.backgroundColor = activeColorProp;
            break;
    }

    let formulaBar = document.querySelector(".formula-bar");
    formulaBar.value = cellProp.formula;
    cell.innerText = cellProp.value;
}

function getActiveCellAndProps(address){
    let [rowId, colId] = decodeRowIdColId(address);

    //Access cell and storage object
    let cell = document.querySelector(`.cell[rowId="${rowId}"][colId="${colId}"]`);
    let cellProp = sheetDB[rowId][colId];
    return [cell, cellProp];
}

function decodeRowIdColId(address){
    // address -> A1
    let rowId = Number(address.slice(1) - 1); //"1" -> 0
    let colId = Number(address.charCodeAt(0)) - 65; // "A" -> 0
    return [rowId, colId];
}






async function traceCyclicPath(cycleResponse){
  let [source_row, source_col] = cycleResponse;
  let visited = [], dfsVisited = [];

  for(let i = 0; i<rows; i++){
      let visitedRow = [], dfsVisitedRow = [];
      for(let j = 0; j<cols; j++){
          visitedRow.push(false);
          dfsVisitedRow.push(false);
      }
      visited.push(visitedRow);
      dfsVisited.push(dfsVisitedRow);
  }
  let response = await dfs(source_row, source_col, visited, dfsVisited);
  return Promise.resolve(response);

}

function colorPromise(){
  return new Promise((resolve, reject) => {
      setTimeout(() => {
          resolve();
      }, 1000)
  })
}

async function dfs(row, col, visited, dfsVisited){
  visited[row][col] = true;
  dfsVisited[row][col] = true;

  let cell = document.querySelector(`.cell[rowId="${row}"][colId="${col}"]`);
  cell.style.backgroundColor = "lightblue";
  await colorPromise(); // wait for 1 sec

  for(let children = 0; children < graphComponentMatrix[row][col].length; children++){
      let [rowId, colId] = graphComponentMatrix[row][col][children];
      if(!visited[rowId][colId]){
          let response = await dfs(rowId, colId, visited, dfsVisited);
          if(response) {
              cell.style.backgroundColor = "transparent";
              await colorPromise();
              return Promise.resolve(true);
          }
      }
      else if(dfsVisited[rowId][colId]) {
          let cyclicCell = document.querySelector(`.cell[rowId="${rowId}"][colId="${colId}"]`);
          cyclicCell.style.backgroundColor = "lightsalmon";
          await colorPromise();

          cyclicCell.style.backgroundColor = "transparent";
          await colorPromise();

          cell.style.backgroundColor = "transparent";
          await colorPromise();

          return Promise.resolve(true);
      }
  }

  dfsVisited[row][col] = false;
  return Promise.resolve(false);
}









for(let i = 0; i<rows; i++){
  for(let j = 0; j<cols; j++){
      let cell = document.querySelector(`.cell[rowId="${i}"][colId="${j}"]`);
      cell.addEventListener("blur", (e) => {
          let address = addressBar.value;
          let [activeCell, cellProp] = getActiveCellAndProps(address);
          let enteredData = activeCell.innerText;

          if(enteredData === cellProp.value) return;

          cellProp.value = enteredData;

          removeChildFromParent(cellProp.formula);
          cellProp.formula = "";
          updateChildrenCells(address);
      })
  }
}

let formulaBar =  document.querySelector('.formula-bar');
formulaBar.addEventListener('keydown', async (e) => {
  let inputFormula = formulaBar.value;
  if(e.key === 'Enter' && inputFormula){
      //check if formula is changed
      let address = addressBar.value;
      let [cell, cellProps] = getActiveCellAndProps(address);
      if(inputFormula !== cellProps.formula) removeChildFromParent(inputFormula);

      addChildToGraphComponent(inputFormula, address);
      // check cyclic formula
      let cycleResponse = isGraphCycle();
      if(cycleResponse){
          let response = confirm("Your formula is cyclic ! Do you want to trace your path ?");
          while(response){
              // keep on tracing
              await traceCyclicPath(cycleResponse);
              response = confirm("Your formula is cyclic ! Do you want to trace your path ?")
          }
          removeChildFromGraphComponent(inputFormula);
          return;
      }

      let evaluatedValue = formulaEvaluator(inputFormula);
      setCellUIAndCellProp(evaluatedValue, inputFormula, address);
      addChildToParent(inputFormula);
      updateChildrenCells(address);
  }
})

function addChildToGraphComponent(formula, childAddress){
  let [child_rowId, child_colId] = decodeRowIdColId(childAddress);
  let encodedFormula = formula.split(" ");
  for(let i = 0; i<encodedFormula.length; i++){
      let asciiValue = encodedFormula[i].charCodeAt(0);
      if(asciiValue >= 65 && asciiValue <= 90){
          let [par_rowId, par_colId] = decodeRowIdColId(encodedFormula[i]);
          graphComponentMatrix[par_rowId][par_colId].push([child_rowId, child_colId]);
      }
  }
}

function removeChildFromGraphComponent(formula){
  let encodedFormula = formula.split(" ");
  for(let i = 0; i<encodedFormula.length; i++){
      let asciiValue = encodedFormula[i].charCodeAt(0);
      if(asciiValue >= 65 && asciiValue <= 90){
          let [par_rowId, par_colId] = decodeRowIdColId(encodedFormula[i]);
          graphComponentMatrix[par_rowId][par_colId].pop();
      }
  }
}

function updateChildrenCells(parentAddress){
  let [parentCell, parentCellProp] = getActiveCellAndProps(parentAddress);
  let childrens = parentCellProp.children;

  for(let i = 0; i<childrens.length; i++){
      let childAddress = childrens[i];
      let [childCell, childCellProp] = getActiveCellAndProps(childAddress);
      let childFormula = childCellProp.formula;
      let evaluatedValue = formulaEvaluator(childFormula);
      setCellUIAndCellProp(evaluatedValue, childFormula, childAddress);
      updateChildrenCells(childAddress);
  }
}

function addChildToParent(formula){
  let encodedFormula = formula.split(" ");
  let childAddress = addressBar.value;
  for(let i = 0; i<encodedFormula.length; i++){
      let firstCharASCII = encodedFormula[i].charCodeAt(0);
      if(firstCharASCII >= 65 && firstCharASCII <= 90){
          let [parentCell, parentCellProp] = getActiveCellAndProps(encodedFormula[i]);
          parentCellProp.children.push(childAddress);
      }
  }
}

function removeChildFromParent(formula){
  let encodedFormula = formula.split(" ");
  let childAddress = addressBar.value;
  for(let i = 0; i<encodedFormula.length; i++){
      let firstCharASCII = encodedFormula[i].charCodeAt(0);
      if(firstCharASCII >= 65 && firstCharASCII <= 90){
          let [parentCell, parentCellProp] = getActiveCellAndProps(encodedFormula[i]);
          parentCellProp.children = parentCellProp.children.filter((child) => child !== childAddress);
      }
  }
}

function formulaEvaluator(formula){
  let encodedFormula = formula.split(" ");
  for(let i = 0; i<encodedFormula.length; i++){
      let firstCharASCII = encodedFormula[i].charCodeAt(0);
      if(firstCharASCII >= 65 && firstCharASCII <= 90){
          let [cell, cellProp] = getActiveCellAndProps(encodedFormula[i]);
          encodedFormula[i] = cellProp.value;
      }
  }
  let decodedFormula = encodedFormula.join(" ");
  return eval(decodedFormula);
}

function setCellUIAndCellProp(evaluatedValue, formula, address){
  let [cell, cellProp] = getActiveCellAndProps(address);
  cell.innerText = evaluatedValue;
  cellProp.value = evaluatedValue;
  cellProp.formula = formula;
}












let ctrlKey;

document.addEventListener("keydown", (e) => {
    ctrlKey = e.ctrlKey;
})
document.addEventListener("keyup", (e) => {
    ctrlKey = e.ctrlKey;
})

for(let i = 0; i<rows; i++){
    for(let j = 0; j<cols; j++){
        let cell = document.querySelector(`.cell[rowId="${i}"][colId="${j}"]`);
        handleSelectedCells(cell);
    }
}

let rangeStorage = [];
let copyBtn = document.querySelector(".copy");
let pasteBtn = document.querySelector(".paste");
let cutBtn = document.querySelector(".cut");

let copyData = [];
copyBtn.addEventListener('click', (e) => {
    if(rangeStorage.length === 0) {
        alert("Select at least one cells");
        return;
    }
    copyData = [];

    if(rangeStorage.length === 1){
        let rid = rangeStorage[0][0], cid = rangeStorage[0][1];
        copyData.push([sheetDB[rid][cid]]);
    }
    else{
        let {start_rid, start_cid, end_rid, end_cid} = getStartEndRange();

        for(let i = start_rid; i <= end_rid; i++) {
            let rowData = [];
            for (let j = start_cid; j <= end_cid; j++) {
                rowData.push(sheetDB[i][j]);
            }
            copyData.push(rowData);
        }
    }
    console.log(copyData);

    defaultSelectedCellsUI();
    rangeStorage = [];
});

pasteBtn.addEventListener('click', (e) => {
    if(copyData.length === 0) {
        alert("Nothing to copy");
        return;
    }

    let address = addressBar.value;
    let [start_rowId, start_colId] = decodeRowIdColId(address);

    for(let i = 0; i < copyData.length; i++){
        for(let j = 0; j < copyData[0].length; j++){
            let rid = i + start_rowId, cid = j + start_colId;
            updateCellProps(rid, cid, copyData[i][j]);
        }
    }
    copyData = [];
})

cutBtn.addEventListener('click', (e) => {
    if(rangeStorage.length === 0) {
        alert("Select at least one cells");
        return;
    }

    copyData = [];

    if(rangeStorage.length === 1){
        let rid = rangeStorage[0][0], cid = rangeStorage[0][1];
        copyData.push([sheetDB[rid][cid]]);
        updateCellProps(rid, cid, defaultCellProps());
    }
    else{
        let {start_rid, start_cid, end_rid, end_cid} = getStartEndRange();

        for(let i = start_rid; i <= end_rid; i++) {
            let rowData = [];
            for (let j = start_cid; j <= end_cid; j++) {
                rowData.push(sheetDB[i][j]);
                updateCellProps(i, j, defaultCellProps());
            }
            copyData.push(rowData);
        }
    }

    defaultSelectedCellsUI();
    rangeStorage = [];
})

function updateCellProps(rowId, colId, newCellProps){
    let cell = document.querySelector(`.cell[rowId="${rowId}"][colId="${colId}"]`);
    if(!cell) {
        console.log("Cell out of bound");
        return;
    }
    sheetDB[rowId][colId] = newCellProps;
    updateCellPropsUI(cell, newCellProps);
}

function handleSelectedCells(cell){
    cell.addEventListener("click", (e) => {
        if(!ctrlKey) return;
        if(rangeStorage.length >= 2) {
            defaultSelectedCellsUI();
            rangeStorage = [];
        }

        cell.style.border = "2px dashed #44a6c6"

        let rowId = Number(cell.getAttribute('rowId'));
        let colId = Number(cell.getAttribute('colId'));

        rangeStorage.push([rowId, colId]);

        if(rangeStorage.length === 2){
            let {start_rid, start_cid, end_rid, end_cid} = getStartEndRange();

            for(let i = start_rid; i <= end_rid; i++){
                for(let j = start_cid; j <= end_cid; j++){
                    let cell = document.querySelector(`.cell[rowId="${i}"][colId="${j}"]`);
                    cell.style.border = "1px solid #dfe4ea";

                    if(i === start_rid) cell.style.borderTop = "2px dashed #44a6c6";
                    if(i === end_rid) cell.style.borderBottom = "2px dashed #44a6c6";
                    if(j === start_cid) cell.style.borderLeft = "2px dashed #44a6c6";
                    if(j === end_cid) cell.style.borderRight = "2px dashed #44a6c6";
                }
            }
        }
    })
}

function defaultSelectedCellsUI(){
    let {start_rid, start_cid, end_rid, end_cid} = getStartEndRange();

    for(let i = start_rid; i <= end_rid; i++){
        for(let j = start_cid; j <= end_cid; j++){
            let cell = document.querySelector(`.cell[rowId="${i}"][colId="${j}"]`);
            cell.style.border = "1px solid #dfe4ea";
        }
    }
}

function getStartEndRange(){
    let start_rid,start_cid, end_rid, end_cid;
    if(rangeStorage.length === 1){
        start_rid = rangeStorage[0][0];
        start_cid = rangeStorage[0][1];
        end_rid = rangeStorage[0][0];
        end_cid = rangeStorage[0][1];
    }
    else {
        start_rid = Math.min(rangeStorage[0][0], rangeStorage[1][0]);
        start_cid = Math.min(rangeStorage[0][1], rangeStorage[1][1]);
        end_rid = Math.max(rangeStorage[0][0], rangeStorage[1][0]);
        end_cid = Math.max(rangeStorage[0][1], rangeStorage[1][1]);
    }
    return {start_rid, start_cid, end_rid, end_cid}
}




let downloadBtn = document.querySelector(".download");
let uploadBtn = document.querySelector(".upload");

downloadBtn.addEventListener("click", (e) => {
    let jsonData = JSON.stringify([sheetDB, graphComponentMatrix]);
    let file = new Blob([jsonData], {type: "application/json"});
    let a = document.createElement('a');
    a.href = URL.createObjectURL(file);
    a.download = "SheetData.json";
    a.click();
})

uploadBtn.addEventListener("click", (e) => {
    let input = document.createElement("input");
    input.setAttribute("type", "file");
    input.click();

    input.addEventListener("change", (e) => {
        let fr = new FileReader();
        let files = input.files;
        let fileObject = files[0];

        fr.readAsText(fileObject);
        fr.addEventListener("load", (e) => {
            let sheetData = JSON.parse(fr.result);
            sheetAddBtn.click();
            sheetDB = sheetData[0];
            graphComponentMatrix = sheetData[1];
            allSheetsDB[allSheetsDB.length - 1] = sheetDB;
            allGraphComponentMatrix[allGraphComponentMatrix.length - 1] = graphComponentMatrix;
            handleSheetProperties();
        })
    })
})