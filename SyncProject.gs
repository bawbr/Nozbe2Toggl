var vg_tokenOfNozbe;
var vg_tokenOfToggl;
var vg_wipOfToggl;

function myFunction() {

  defineGlobalVariables();

  updateProjectListforNew();
  updateProjectListforCompleted();

  editProjectList();

  updateTogglforNew();
  updateTogglforCompleted();
  
}

function defineGlobalVariables(){
  var ss = SpreadsheetApp.openById(SpreadsheetApp.getActiveSpreadsheet().getId());
  var sheet = ss.getSheetByName("AuthSettings");
  var range_tokenOfNozbe = sheet.getRange(1,2,1,2);
  var range_tokenOfToggl = sheet.getRange(2,2,2,2);
  var range_wipOfToggl = sheet.getRange(3,2,3,2);
  
  vg_tokenOfNozbe = range_tokenOfNozbe.getValue();
  vg_tokenOfToggl = range_tokenOfToggl.getValue();
  vg_wipOfToggl = range_wipOfToggl.getValue();
}


function updateTogglforNew(){
  var ss = SpreadsheetApp.openById(SpreadsheetApp.getActiveSpreadsheet().getId());
  var sheet = ss.getSheetByName("ProjectList");
  var sheet2 = ss.getSheetByName("LabelList");
  var lastRow = sheet.getLastRow();
  var lastRow2 = sheet2.getLastRow();
  
  var range_FlagNew = sheet.getRange(2,17,lastRow-1);
  var val_FlagNew = range_FlagNew.getValues();

  var range_TogglProjectName = sheet.getRange(2,4,lastRow-1);
  var val_TogglProjectName = range_TogglProjectName.getValues();
  
  var range_TogglProjectIDs = sheet.getRange(2,15,lastRow-1);
  var val_TogglProjectIDs = range_TogglProjectIDs.getValues();
  
  var range_NozbeTag = sheet.getRange(2,13, lastRow-1);
  var val_NozbeTag = range_NozbeTag.getValues();
  
  var range_NozbeTagMaster = sheet2.getRange(2,1,lastRow2-1);
  var val_NozbeTagMaster = range_NozbeTagMaster.getValues();
  
  var range_TogglClientMaster = sheet2.getRange(2,2,lastRow2-1);
  var val_TogglClientMaster = range_TogglClientMaster.getValues(); 
  
  for (var i in val_FlagNew){
    if ( val_FlagNew[i][0] == true && val_TogglProjectIDs[i][0] == '' && val_TogglProjectName[i][0] != ''){
      if ( val_NozbeTag[i][0] == ''){    
        val_TogglProjectIDs[i][0] = toggl(val_TogglProjectName[i][0]).data.id;
      }
      
      if ( val_NozbeTag[i][0] != '' ){
        for (var j in val_NozbeTagMaster ){
          if (val_NozbeTagMaster[j][0] == val_NozbeTag[i][0] ){
            val_TogglProjectIDs[i][0] = togglwcl(val_TogglProjectName[i][0],val_TogglClientMaster[j][0]).data.id;
          }
        }
      }
      
      val_FlagNew[i][0] = false;
    }    
  }
  
  range_FlagNew.setValues(val_FlagNew);
  range_TogglProjectIDs.setValues(val_TogglProjectIDs);
  
}

function createFileInGoogleDrive(){
}


function updateTogglforCompleted(){
  var ss = SpreadsheetApp.openById(SpreadsheetApp.getActiveSpreadsheet().getId());
  var sheet = ss.getSheetByName("ProjectList");
  var cell = sheet.getRange('n1');

  var range_TogglProjectIDs = sheet.getRange(2,15,sheet.getLastRow()-1);
  var val_TogglProjectIDs = range_TogglProjectIDs.getValues();

  var range_NozbeCompleted = sheet.getRange(2,14,sheet.getLastRow()-1);
  var val_NozbeCompleted = range_NozbeCompleted.getValues();

  var range_TogglCompleted = sheet.getRange(2,16,sheet.getLastRow()-1);
  var val_TogglCompleted = range_TogglCompleted.getValues();
  
  Logger.log(val_NozbeCompleted);
  Logger.log(val_TogglCompleted);
  Logger.log(val_TogglProjectIDs);
  
  for ( var i in val_NozbeCompleted ){
    if ( val_TogglProjectIDs[i][0] != '' && val_TogglCompleted[i][0] != true && val_NozbeCompleted[i][0] == true ){
      Logger.log(i);
      var ret = archiveTogglProject(val_TogglProjectIDs[i][0]);
      if (ret != null){
        if (ret.data.id != undefined ){
          val_TogglCompleted[i][0] = true;
        }
      }
    }
  }
  
  range_TogglCompleted.setValues(val_TogglCompleted);

}


function updateProjectListforNew(){
  var ss = SpreadsheetApp.openById(SpreadsheetApp.getActiveSpreadsheet().getId());
  var sheet = ss.getSheetByName("ProjectList");
  var cell = sheet.getRange('a1');
  var newProjects = getNewProject();
  var lastRow = sheet.getLastRow();
  for ( var i in newProjects ){
    var cell0 = cell.offset(lastRow,0);
    cell0.offset(0,4).setValue(newProjects[i].id);
    cell0.offset(0,5).setValue(newProjects[i].name);
    cell0.offset(0,6).setValue(newProjects[i].body);
    cell0.offset(0,7).setValue(newProjects[i]._created_at);
    cell0.offset(0,8).setValue(newProjects[i].guid);
    cell0.offset(0,9).setValue(newProjects[i]._share_people);
    cell0.offset(0,10).setValue(newProjects[i].description);
    cell0.offset(0,11).setValue(newProjects[i]._has_completed);
    cell0.offset(0,12).setValue(newProjects[i].tags);
    lastRow++; 
  }
}

function updateProjectListforCompleted(){
  var ss = SpreadsheetApp.openById(SpreadsheetApp.getActiveSpreadsheet().getId());
  var sheet = ss.getSheetByName("ProjectList");
  var cell = sheet.getRange('n1');

  var range_ProjectIDs = sheet.getRange(2,5,sheet.getLastRow()-1);
  var val_ProjectIDs = range_ProjectIDs.getValues();

  var range_completed = sheet.getRange(2,14,sheet.getLastRow()-1);
  var val_completed = range_completed.getValues();
  
  var openProjectIDs = getOpenProjectIDs();
  
  for ( var i in val_ProjectIDs ){
    if(! openProjectIDs.some(function(v){return v == val_ProjectIDs[i] } ) ){
        val_completed[i][0] = true;
    } 
  }
  

  range_completed.setValues(val_completed);

}

function getOpenProjectIDs(){
  var openProjects = getOpenProjects();
  var openProjectIDs = new Array;
  
  for (var i in openProjects ){
    openProjectIDs.push(openProjects[i].id);
  }
  return openProjectIDs;
}

function getOpenProjects(){
  var NozbeProjects = getNozbeProjects();
  var openProjects = new Array;

  for ( var i in NozbeProjects){
       var openProject = {
        'id' : NozbeProjects[i].id,
        'name' : NozbeProjects[i].name,
        'body' : NozbeProjects[i].body,
        '_created_at' : NozbeProjects[i]._created_at,
        'guid' : NozbeProjects[i].guid,
        '_share_people' : NozbeProjects[i]._share_people,
        'description' : NozbeProjects[i].description,
        '_has_completed' : NozbeProjects[i]._has_completed,
        'tags' : NozbeProjects[i].tags
       };
      openProjects.push(openProject);
  }
  
  return openProjects;

}

function getCompletedProjectIDs(){
  var completedProjects = getCompletedProjects();
  var completedProjectIDs = new Array;
  
  for (var i in completedProjects ){
    completedProjectIDs.push(completedProjects[i].id);
  }
  return completedProjectIDs;
}
  
function getCompletedProjects(){
  var NozbeProjects = getNozbeProjects();
  var completedProjects = new Array;

  for ( var i in NozbeProjects){
    if ( NozbeProjects[i]._has_completed == true) {
      var completedProject = {
        'id' : NozbeProjects[i].id,
        'name' : NozbeProjects[i].name,
        'body' : NozbeProjects[i].body,
        '_created_at' : NozbeProjects[i]._created_at,
        'guid' : NozbeProjects[i].guid,
        '_share_people' : NozbeProjects[i]._share_people,
        'description' : NozbeProjects[i].description,
        '_has_completed' : NozbeProjects[i]._has_completed,
        'tags' : NozbeProjects[i].tags
      };
      completedProjects.push(completedProject);
    }
  }
  
  return completedProjects;
}  


function getNewProject(){
  var ss = SpreadsheetApp.openById(SpreadsheetApp.getActiveSpreadsheet().getId());
  var sheet = ss.getSheetByName("ProjectList");
  var cell_id = sheet.getRange('E:E').getValues();
  var NozbeProjects = getNozbeProjects();
  var newProjects = new Array;
  
  for ( var i in NozbeProjects){
    if ( ! cell_id.some(function(v){return v == NozbeProjects[i].id}) ){
      var newProject = {
        'id' : NozbeProjects[i].id,
        'name' : NozbeProjects[i].name,
        'body' : NozbeProjects[i].body,
        '_created_at' : NozbeProjects[i]._created_at,
        'guid' : NozbeProjects[i].guid,
        '_share_people' : NozbeProjects[i]._share_people,
        'description' : NozbeProjects[i].description,
        '_has_completed' : NozbeProjects[i]._has_completed,
        'tags' : NozbeProjects[i].tags
      };
      newProjects.push(newProject);
    }
  } 
  return newProjects;
  
}

function editProjectList(){
  var ss = SpreadsheetApp.openById(SpreadsheetApp.getActiveSpreadsheet().getId());
  var sheet = ss.getSheetByName("ProjectList");
  var lastRow = sheet.getLastRow();
  
  var range_PlainProjectName = sheet.getRange(2,1,lastRow-1);
  var val_PlainProjectName = range_PlainProjectName.getValues();
   
  var range_NozbeProjectName = sheet.getRange(2,6,lastRow-1);
  var val_NozbeProjectName = range_NozbeProjectName.getValues();
 
  var range_CreatedAt = sheet.getRange(2,3,lastRow-1);
  var val_CreatedAt = range_CreatedAt.getValues();
  
  var range_TogglProjectName = sheet.getRange(2,4,lastRow-1);
  var val_TogglProjectName = range_TogglProjectName.getValues();
  
  var range_TogglProjectID = sheet.getRange(2,15,lastRow-1);
  var val_range_TogglProjectID = range_TogglProjectID.getValues();
  
  var range_NozbeCreatedAt = sheet.getRange(2,8,lastRow-1);
  var val_NozbeCreatedAt = range_NozbeCreatedAt.getValues();
  
  var range_FlagNew = sheet.getRange(2,17,lastRow-1);  
  var val_FlagNew = range_FlagNew.getValues();
    
  for (var i in val_NozbeProjectName ){
    
    Logger.log(val_PlainProjectName[i][0]);
    
    if ( val_PlainProjectName[i][0] == '' ){
      if ( val_NozbeProjectName[i][0] != '' ){
        val_PlainProjectName[i][0] = val_NozbeProjectName[i][0];
        val_FlagNew[i][0] = true;
      }
    }
    
    if ( val_CreatedAt[i][0] == '' ){
      if ( val_NozbeCreatedAt[i][0] != '' && val_NozbeCreatedAt[i][0] != 'undefined' ) {
        val_CreatedAt[i][0] = val_NozbeCreatedAt[i][0];
        val_FlagNew[i][0] = true;
      }
    }
    
    if ( val_TogglProjectName[i][0] == '' ){
      val_TogglProjectName[i][0] = val_NozbeProjectName[i][0];
      if ( val_NozbeCreatedAt[i][0] != '' ){
        val_TogglProjectName[i][0] = Utilities.formatDate(new Date(val_CreatedAt[i][0]), 'JST', 'yyyyMM') + '_' + val_TogglProjectName[i][0];
      }
      val_FlagNew[i][0] = true;
    }    
  }
  
  range_PlainProjectName.setValues(val_PlainProjectName);
  range_CreatedAt.setValues(val_CreatedAt);
  range_TogglProjectName.setValues(val_TogglProjectName);
  range_FlagNew.setValues(val_FlagNew);

}


function toggl(project){
  
  var retJSON = newTogglProject(project);
  
  if (retJSON){
    try {
      var ret = JSON.parse(retJSON);
    } catch(e) {
      var id = {
        'id' : 'error'
      };
      
      var ret = {
        data : id
      };
    }
    return ret;
  }
}

function togglwcl(project,cid){
  var retJSON = newTogglProjectwithClient(project,cid);
  
  if (retJSON){
    try {
      var ret = JSON.parse(retJSON);
    } catch(e) {
      var id = {
        'id' : 'error'
      };
      
      var ret = {
        data : id
      };
    }
    return ret;
  }
}

function newTogglProjectwithClient(project,cid){
  var url = 'https://www.toggl.com/api/v8/projects';
  var token = vg_tokenOfToggl;
  var token64 = Utilities.base64Encode(token);
  var auth = 'Basic ' + token64;
  var name = project;
  var wip = vg_wipOfToggl;
  var client = new Number(cid);
  var is_private = true;
  var billable = true;
  var method = 'post';
  var contenttype = 'Content-Type: application/json';
  var muteHttpExceptions = true;
  
  var data = {
    'Authorization'      : auth
  };
  
  var projectdata = {
    'name'           : name,
    'wip'            : wip,
    'active'         : true,
    'is_private'     : true,
    'cid'            : client
  };
  
  var payload = {
    'project' : projectdata
  };
  
  
  var params = {
    'method' : method,
    'contentType' : contenttype,
    'headers' : data,
    'payload' :  JSON.stringify(payload),
    'muteHttpExceptions' : muteHttpExceptions
  };

  
  var response = UrlFetchApp.fetch(url, params);
  Logger.log(response);
  return response;
}

function newTogglProject(project){
  var url = 'https://www.toggl.com/api/v8/projects';
  var token = vg_tokenOfToggl;
  var token64 = Utilities.base64Encode(token);
  var auth = 'Basic ' + token64;
  var name = project;
  var wip = vg_wipOfToggl;
  var is_private = true;
  var billable = true;
  var method = 'post';
  var contenttype = 'Content-Type: application/json';
  var muteHttpExceptions = true;
  
  var data = {
    'Authorization'      : auth
  };
  
  var projectdata = {
    'name'           : name,
    'wip'            : wip,
    'active'         : true,
    'is_private'     : true
  };
  
  var payload = {
    'project' : projectdata
  };
  
  
  var params = {
    'method' : method,
    'contentType' : contenttype,
    'headers' : data,
    'payload' :  JSON.stringify(payload),
    'muteHttpExceptions' : muteHttpExceptions
  };

  
  var response = UrlFetchApp.fetch(url, params);
  Logger.log(response);
  return response;
}


function getNozbeProjects(){
  var url        = 'https://api.nozbe.com:3000/list';
  var access_token = vg_tokenOfNozbe;
  var method = 'get';
  var type = 'project';
  
  var headers = {
    'Authorization' : access_token,
  };
    
  url = url + '?type=' + type;
  
  var params = {
    'method' : method,
    'headers' : headers,
  };
  
  var JSONresponse = UrlFetchApp.fetch(url, params);

  var response = JSON.parse(JSONresponse);  
    
  return response;  
  
}

function getTogglProjects(){
  var url = 'https://www.toggl.com/api/v8/workspaces/vg_wipOfToggl/projects';
  var token = vg_tokenOfToggl;
  var token64 = Utilities.base64Encode(token);
  var auth = 'Basic ' + token64;
  var wip = vg_wipOfToggl;
  var is_private = true;
  var billable = true;
  var method = 'get';
  var contenttype = 'Content-Type: application/json';
  var muteHttpExceptions = true;
  
  var data = {
    'Authorization'      : auth
  };
    
  var params = {
    'method' : method,
    'contentType' : contenttype,
    'headers' : data,
  };

  
  var response = UrlFetchApp.fetch(url, params);
  var ret = JSON.parse(response);
  for ( var i in ret){
    Logger.log('\t' + ret[i].id + '\t' + ret[i].name );
  }

}

  
function getTogglClients(){
  var url = 'https://www.toggl.com/api/v8/workspaces/vg_wipOfToggl/clients';
  var token = vg_tokenOfToggl;
  var token64 = Utilities.base64Encode(token);
  var auth = 'Basic ' + token64;
  var wip = vg_wipOfToggl;
  var is_private = true;
  var billable = true;
  var method = 'get';
  var contenttype = 'Content-Type: application/json';
  var muteHttpExceptions = true;
  
  var data = {
    'Authorization'      : auth
  };
    
  var params = {
    'method' : method,
    'contentType' : contenttype,
    'headers' : data,
  };

  
  var response = UrlFetchApp.fetch(url, params);
  var ret = JSON.parse(response);
  for ( var i in ret){
    Logger.log('\t' + ret[i].id + '\t' + ret[i].name );
  }
  
}

function archiveTogglProject(ProjectID){
  try {
    var ret = JSON.parse(archiveTogglProjectJSON(ProjectID));
  } catch (e) {
    var ret = null;
  }
  Logger.log(ret);
  return ret;
}

 
function archiveTogglProjectJSON(ProjectID){
    
  var url = 'https://www.toggl.com/api/v8/projects';
  var token = vg_tokenOfToggl;
  var token64 = Utilities.base64Encode(token);
  var auth = 'Basic ' + token64;
  var wip = vg_wipOfToggl;
  var active = false;
  var method = 'put';
  var contenttype = 'Content-Type: application/json';
  var muteHttpExceptions = true;
  var url = url + '/' + ProjectID;
  var headers = {
    'Authorization'      : auth
  };
  
  var projectdata = {
    'active'         : active
  };
  
  var payload = {
    'project' : projectdata
  };
  
  var params = {
    'method' : method,
    'contentType' : contenttype,
    'headers' : headers,
    'payload' :  JSON.stringify(payload),
    'muteHttpExceptions' : muteHttpExceptions
  };

  Logger.log(params);
  Logger.log(url);
  
  var response = UrlFetchApp.fetch(url, params);
  Logger.log(response);
  return response;
   
}


