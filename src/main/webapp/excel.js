"use strict";
/*
 * Javascript for managing Excel spreadsheet uploads
 * 
 * Copyright 2014, Chiral Software, Inc.
 */

function activateExcel() {
    var excelTable = document.getElementById("lineSelectorTr");
    var selectors = excelTable.getElementsByTagName("select");
    for(var i = 0 ; i < selectors.length; i++) {
        selectors[i].onchange = selectorChanged;
        selectors[i].selectedIndex = i + 1;
    }
    enableSubmit();
}

function selectorChanged() {
    if(this.selectedIndex === 0) {
        // we are in an unselected state
        // IF all the others are also unselected,
        // then disable the submit button
        enableSubmit();
        return;
    }
    var selectedValue = this.options[this.selectedIndex].value;
    console.log("my selected value is: " + selectedValue);
    // we are in a selected state so make sure that no other selector is the same column name
    var excelTable = document.getElementById("lineSelectorTr");
    var selectors = excelTable.getElementsByTagName("select");
    for(var i = 0 ; i < selectors.length; i++) {
        if(selectors[i] != this) {
            if(selectors[i].options[selectors[i].selectedIndex].value === selectedValue) {
                selectors[i].selectedIndex = 0;
            }
        }
    }
    enableSubmit();
}

function enableSubmit() {
    // go through all the selectors, and make sure that we have 
    // values for name and phone, at a minimum
    var excelTable = document.getElementById("excelTable");
    var selectors = excelTable.getElementsByTagName("select");
    var submitButton = document.getElementById("saveExcel");
    for(var i = 0 ; i < selectors.length; i++) {
        var selectedValue = selectors[i].options[selectors[i].selectedIndex].value;
        if(selectedValue !== "-") {
            submitButton.disabled = false; 
            return;
        }
    }
    submitButton.disabled = true;
}