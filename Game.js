// HELLO THIS IS MY AWESOME GAME CODE! 
// Okay so first of all, we need to like actually get the game sheet
function getGameSheet(){
    return SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
}
// If your controller's not working, run the function below on AppsScript
function RunThisFunctionFirst() {
    const sheet = getGameSheet();
    const spreadsheet = sheet.getParent()
    sheet.getRange("H16").insertCheckboxes();
    sheet.getRange("B16").setValue("click 2 start ->");
    for (const trigger of ScriptApp.getProjectTriggers()) {
        if (trigger.getHandlerFunction() === "controllerEdited") {
            ScriptApp.deleteTrigger(trigger);
        }
    }

    ScriptApp.newTrigger("controllerEdited")
        .forSpreadsheet(spreadsheet)
        .onEdit()
        .create();

    console.log("Controller connected to: " + spreadsheet.getName());
    spreadsheet.getSheets()[0].getRange("B15").setValue("connektion accomplished");
}

// lets make the board
// consts are used to make variables that will remain the same throughout the entire project 
const board_int_row = 3; // edit these to change where the board starts
const board_int_col = 3; // < this one too
const board_size = 12;  // this one changes the board size
const board_row_size = board_size;  // I guess if you want to change the size of rows or columns you could mess with these?
const board_col_size = board_size; 
const boarder_int_row = board_int_row - 1;  // I hope this will make it go to the upper right corner?
const boarder_int_col = board_int_col - 1;
const boarder_row_size = board_row_size + 2; // this is the size of outline, the large the number the thiccer the outline. 
const boarder_col_size = board_row_size + 2;
// I need to make a variable, apparently a variable in JAVASCRIPT is made using "let" and instead of the libraries that python has, we use ARRAYS
// The snakes pixels are kinda an array ?
let snake = [
    [5,5],
    [5,4],
    [5,3]
]; 

// we can draw the outline and the board (outline goes first because it fills the back first)
function doodleGame() {
    const sheet = getGameSheet();

    const outline = getGameSheet().getRange(
        boarder_int_row,
        boarder_int_col,
        boarder_col_size,
        boarder_row_size
    )
    const board = sheet.getRange(
        board_int_row,
        board_int_col, 
        board_row_size,
        board_col_size
    );
    // This section defines the pixel heights 
    sheet.setColumnWidths(boarder_int_col,boarder_col_size, 30);
    sheet.setRowHeights(boarder_int_row,boarder_row_size, 30);
    // Here it defines the colors
    outline.setBackground("red")
    board.setBackground("black");
}

// this reads the controller 
function readCtrl() {
    const sheet = getGameSheet();
    // we need to make sure it's not case sensitive :D
    const key = String(sheet.getRange(1, 8).getValue())
        .trim()
        .toUpperCase();
    console.log(key);
    if (key === "W") return "U";
    if (key === "A") return "L";
    if (key === "S") return "D";
    if (key === "D") return "R";
    return null;
}

// This checks the game status
function gameOn() {
    const sheet = getGameSheet(); 
    let gameStat = sheet.getRange(1,12).getValue();
    return gameStat
}
// everything works :D
function doodleController() {
    const sheet = getGameSheet();
    sheet.getRange(1,8).setBackground("blue").setFontColor("yellow").setHorizontalAlignment("center").setFontFamily("DotGothic16");
}


// This spawns the snake
function doodleSnake() {
    const sheet =  getGameSheet();
    for (let i=0; i < snake.length; i++){
        let snakeC = board_int_col + snake[i][1] - 1;
        let snakeR = board_int_row + snake[i][0] - 1;
        let specficell = sheet.getRange(snakeR,snakeC);
        specficell.setBackground("yellow");
    }
}
// this is the title! 
function RunThisFunctionToStartTheGame(){
    const sheet = getGameSheet();
    PropertiesService.getScriptProperties().setProperty("gameStarted","false");
    doodleGame();
    const board = sheet.getRange( 
    board_int_row, board_int_col, board_row_size, board_col_size);
    board.breakApart();
    board.clearContent();
    board.merge()
        .setBackground("#012611")
        .setFontColor("#214c0c")
        .setFontSize(24)
        .setFontFamily("DotGothic16")
        .setFontWeight("bold")
        .setHorizontalAlignment("center")
        .setVerticalAlignment("middle")
        .setWrap(true)
        .setValue("WELCOME TO MY AWESOME SNAKE GAME \n\n click the checkbox to start!");
    sheet.getRange("H1").clearContent();
    sheet.getRange("I1").setValue("Click the checkbox 2 start");

}



let kiwi_place = [4,8];
// we need to spawn the kiwi somehow
function doodleKiwi(){
    if (kiwi_place === null) return; 
    const sheet = getGameSheet();
    let kiwiC = board_int_col + kiwi_place[1] -1;
    let kiwiR = board_int_row + kiwi_place[0] -1; 
    const kiwis = sheet.getRange(kiwiR,kiwiC);
    kiwis.setBackground("green");

}
function startGame () {
    const sheet = getGameSheet();
    const storage = PropertiesService.getScriptProperties();
    storage.setProperty("gameStarted", "false");
    const board = sheet.getRange(
        board_int_row,
        board_int_col, 
        board_row_size,
        board_col_size
    );
    board.breakApart();
    board.clearContent();
    resetSnake();
    sheet.getRange("I1").setValue("<- Enter a WASD move into H1").setFontColor("#214c0c").setFontFamily("DotGothic16");
    sheet.getRange("H1").clearContent().activate();
    storage.setProperty("gameStarted", "true");


}

// ok next they need to 

// ok we need a function that explains where the things moving...
function saveSnake() {
    const storage = PropertiesService.getScriptProperties();
    // Im reading the thing FROM script and then turning it into readable stuffs. 
    storage.setProperty("snakeposition", JSON.stringify(snake));
    storage.setProperty("kiwiposition", JSON.stringify(kiwi_place));
}
// This gets the snakes position for reference
function loadSnake() {
    const storage = PropertiesService.getScriptProperties();

    let saved = storage.getProperty("snakeposition");

    if (saved != null) {
        snake = JSON.parse(saved);
    } 
    const siwi = storage.getProperty("kiwiposition");

    if (siwi !== null) {
     kiwi_place = JSON.parse(siwi);
    }
}
// Does the snake hit the wall? 
function outOfBounds(newhead){
    let row = newhead[0];
    let col = newhead[1];

    if (row < 1){
        return true;
    }
    if (row >  board_row_size) {
        return true;
    }
    if (col < 1){
        return true;
    }
    if (col > board_col_size ){
        return true;
    }
    return false;

}
// this spawns the kiwi somewhere random 
function lettherebekiwi() {
    // first I make up a board
    const emptyCells = [];
    for (let row = 1; row <= board_row_size; row++) {
        for (let col = 1; col <= board_col_size; col++) {
            const occupied = snake.some(segment =>
                segment[0] === row && segment[1] === col
            );
            if (!occupied) {
                emptyCells.push([row, col]);
            }
        }
    }
    //win condition
    if (emptyCells.length === 0) {
        getGameSheet().getRange("I1").setValue("You win! Yay :D").setFontColor("#63ba38").setFontFamily("DotGothic16");
        return null;
    }
    // Spawn the kiwi somewhere random
    const index = Math.floor(Math.random() * emptyCells.length);
    return emptyCells[index];
}
// this takes what the controller read and makes it move the array
function moveSnake(where) {
    if (!["R","L","U","D"].includes(where)){
        console.log("Not a direction!")
        return
    }
    let head = snake[0];
    let newhead = [head[0], head[1]];
    if (where === "R") {
    newhead[1] += 1;
    }
    if (where === "L") {
    newhead[1] -= 1;
    }
    if (where === "D"){
    newhead[0] += 1;
    }
    if (where === "U") {
    newhead[0] -= 1;
    }
    if (
    snake.length > 1 &&
    newhead[0] === snake[1][0] &&
    newhead[1] === snake[1][1]
) {
    // making sure kids don't go backwards 
    getGameSheet().getRange("I1").setValue("You can't go backwards").setFontColor("#4c0c41").setFontFamily("DotGothic16");
    return;
}   // making sure you don't hit a wall
    if(outOfBounds(newhead)){
        getGameSheet().getRange("I1").setValue("Game over :<").setFontColor("#4c0c16").setFontFamily("DotGothic16");
        resetSnake();
        return;
    }
    const eating = kiwi_place !== null &&
    newhead[0] === kiwi_place[0] &&
    newhead[1] === kiwi_place[1];

    const bodyLength = eating ? snake.length : snake.length - 1;

    for (let i = 0; i < bodyLength; i++) {
    if (
        newhead[0] === snake[i][0] &&
        newhead[1] === snake[i][1]
    ) {
        getGameSheet().getRange("I1").setValue("game over :(").setFontColor("#4c0c16").setFontFamily("DotGothic16");
        resetSnake();
        return;
    }
}
    snake.unshift(newhead);
    if (
        kiwi_place !== null &&
        newhead[0] === kiwi_place[0] &&
        newhead[1] === kiwi_place[1]
    )
    {console.log("Nomnom");
        kiwi_place = lettherebekiwi();
    } else {
        snake.pop();}
}
// This writes the score
function doodleScore() {
    const sheet = getGameSheet();
    const storage = PropertiesService.getScriptProperties();

    const score = snake.length -3;
    let hiscore = Number(storage.getProperty("hiscore") || 0);
    if(!Number.isFinite(hiscore)){
        hiscore = 0;
    }

    hiscore = Math.max(hiscore, score)
    storage.setProperty("hiscore", String(hiscore));

    sheet.getRange("B1").setValue("Score=" + score)
    sheet.getRange("D1").setValue("High Score=" + hiscore)


}

// RESET the snake so that it doesn't explode :D
function resetSnake() {
    snake = [[5,5],[5,4],[5,3]];
    kiwi_place = [4,8];

    saveSnake(); 
    doodleGame();
    doodleKiwi();
    doodleScore();
    doodleController();
    doodleSnake();
    console.log("Snake reset...");
}
// This checks the things happening in the cell
function controllerEdited(e) {
    if (!e) return;

    const cell = e.range;
    const sheet = cell.getSheet();

    if (sheet.getSheetId() !== getGameSheet().getSheetId()) return;
    if(cell.getA1Notation() === "H16") {
        if(cell.isChecked() === true){
            startGame();
            cell.uncheck();
        } return; }
    if (cell.getA1Notation() !== "H1") return;

    runGame();
    cell.clearContent();
}
// running the game
function runGame() {
    // make sure the game has actually started
    if (
        PropertiesService.getScriptProperties().getProperty("gameStarted") !== "true") {return}
    const where = readCtrl();
    // Since they can type in anything in the cell, this makes sure that it isn't another key! 
    if (where === null){
        console.log("I can't understand your keystroke...")
        return;
    }
    // this remembers the snakes past position 
    loadSnake();
    // makes sure the snake knows where to move next
    if (snake.length === board_row_size * board_col_size){
        getGameSheet().getRange("I1").setValue("YAY YOU WIN!! click continue to play again;)"); 
        return;
    }
    moveSnake(where)
    // saves the snake in its place 
    saveSnake();
    // now we're running the function that makes the game appear
    doodleGame();
    // The snake too
    doodleSnake();
    // and the kiwi
    doodleKiwi();
    // & the score
    doodleScore();
    // and where you're supposed to write WASD TOO! 
    doodleController(); 
}
