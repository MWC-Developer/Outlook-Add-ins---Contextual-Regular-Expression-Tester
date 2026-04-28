// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. see LICENSE in the project root for license information.

Office.initialize = function(reason) {
  console.log("Office.initialize called with reason: " + reason);
  $(document).ready(function(){
    loadEntities();
  });
}

function loadEntities() {
  console.log("Loading entities...");
  if (Office.context.mailbox.item.getSelectedRegExMatches !== undefined) {
    console.log("getSelectedRegExMatches is supported");
    var selectedMatches = Office.context.mailbox.item.getSelectedRegExMatches();
    if (selectedMatches) {
      $("#selected-match").text(JSON.stringify(selectedMatches, null, 2));
      console.log("Selected matches: " + JSON.stringify(selectedMatches, null, 2));
    } else {
      $("#selected-match").text("Selected matches was null");
      console.log("Selected matches was null");
    }
  } else {
    $("#selected-match").text("Method not supported on your client");
    console.log("getSelectedRegExMatches is not supported on your client");
  }

  // Get all matches
  var allMatches = Office.context.mailbox.item.getRegExMatches();
  if (allMatches) {
    $("#all-matches").text(JSON.stringify(allMatches, null, 2));
    console.log("All matches: " + JSON.stringify(allMatches, null, 2));
  } else {
    $("#all-matches").text("All matches was null");
    console.log("All matches was null");
  }
}

function showError(message) {
  $("#error-msg").text(message);
  $("#error").show();
  console.error(message);
}