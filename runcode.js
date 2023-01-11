function processEdit(sheetName, row) {
  //get the last name
  let playerName = 'Storer';
  //check whether it's an in or out
  if (playerName !== null) {
    console.log(playerName);
    const InOutWant = 'In';
    //switch on sheetName
    switch (sheetName) {
      case 'In/Out':
        console.log('case In/Out ' + playerName);
        switch (InOutWant) {
          case 'In':
            console.log('case In');
            //if it's an in, check whether value is Y
            console.log(playerName);
            break;
          case 'Out':
            //get the last name
            //set value to N
            setYNCW(playerName, 'N');
            break;
          default:
            break;
        }
        break;
      case 'Want to Play':
        switch (InOutWant) {
          case 'Want':
            //check if the value is a Y
            const currentYNCW = getYNCW(playerName);
            if (currentYNCW === 'Y') {
              //if so, set value to C
              setYNCW(playerName, 'C');
            } else {
              processWantEdit(playerName);
            }
            break;
          default:
            break;
        }
      default:
        break;
    }
  }
}

processEdit('In/Out', 10);
