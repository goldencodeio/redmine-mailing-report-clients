function insertText() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var rowI = 1;
  var columnI = 1;

  OPTIONS.numbersContracts.forEach(function(contract) {
    var allProjects = APIRequest('projects').projects;
    var pattern = new RegExp('^' + contract + '-[0-9]+-SUP');
    var projects = allProjects.filter(function(project) {
      return pattern.test(project.name);
    });

    projects.forEach(function(proj) {
      Logger.log(proj.name);
    });

    Logger.log('------------------------');
    rowI++;
  });
}
