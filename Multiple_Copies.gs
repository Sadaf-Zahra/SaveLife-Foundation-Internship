function makeMultipleCopies() {
  const TEMPLATE_ID = '1oN9JtyxZfUl4wKrhPCifIRi6lw2Hked1DSwYHsjxwKE'; 
  const FOLDER_ID = 'DRIVE_LINK_WHERE_FILES_ARE_TO_BE_SAVED';  // Your folder ID

  for (let i = 1; i <= 100; i++) {
    const name = 'Slide Copy ' + i;
    const copy = DriveApp.getFileById(TEMPLATE_ID).makeCopy(name, DriveApp.getFolderById(FOLDER_ID));
    Logger.log('Created: ' + name + ' ID: ' + copy.getId());
  }
}
