function build_contribution(projectListSheetName, lookup_member_name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var project_list = ss.getSheetByName(projectListSheetName);
  var lastRow = project_list.getLastRow();
  var total_max_member_count = 6;
  var final_member_contribution = {};
  
  for (var i=2 ; i<=lastRow ; ++i) {
    var cell = project_list.getRange(i,1);
    var project_value = parseFloat(project_list.getRange(i,3).getValues()[0]);
    
    for (var base=5, j=base ; j <base+total_max_member_count ; j+=2) {
      var project_member_name = project_list.getRange(i,j).getValues()[0];
      project_member_name = project_member_name.toString().toLowerCase();
      var project_member_contribution = parseFloat(project_list.getRange(i,j+1).getValues()[0]);
      if (project_member_name != '' && project_member_contribution != '') {
        
        if (!final_member_contribution[project_member_name])
          final_member_contribution[project_member_name] = project_value * project_member_contribution;
        else
          final_member_contribution[project_member_name] += project_value * project_member_contribution;
      }
    }
  }
  lookup_member_name = lookup_member_name.toString().toLowerCase();
  return final_member_contribution[lookup_member_name] ? final_member_contribution[lookup_member_name] : -1;
}
