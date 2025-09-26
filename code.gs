function onOpen() {
    const ui = SpreadsheetApp.getUi();
    const menu = ui.createMenu('Generate Signable Doc');
    menu.addItem('Create New Docs', 'createNewGoogleDocs')
    menu.addToUi();

}

//date function
function theDate(){
var d = new Date();
var curr_day = d.getDate();
var curr_month = d.getMonth() + 1; //Months are zero based
var curr_year = d.getFullYear();

var theDate = curr_month + "-" + curr_day + "-" + curr_year;

return theDate;

}

function createNewGoogleDocs() {
    
    //HF template ie template with header and footer
    const googleDocTemplate = DriveApp.getFileById('TemplateID');
    
    //Doc2Print folder we want to save the files in
    const destinationFolder = DriveApp.getFolderById('FolderToSaveID')
    //name of doc we are saving
    const docCopy = googleDocTemplate.makeCopy("Sign Doc "+ theDate(), destinationFolder);
    const doc = DocumentApp.openById(docCopy.getId())

    //get 'Data' sheet
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Data')

    //header cells 2x5 the second blank cells are just for formatting.
    const cells = [['LAST', 'FIRST', 'DATE', 'CHECK', 'SIGNATURE'],[' ', ' ', ' ', ' ', ' ']];
    
    //sort sheet by where
    //so each person will be grouped by where
    //Updated to be row 6 vs 5, need to freeze the rows
    sheet.setFrozenRows(1);
    sheet.sort(6, true);
    sheet.setFrozenRows(0);
    
    //display values so we get what we see in dates and such
    const rows = sheet.getDataRange().getDisplayValues();
    
    //temp var to setup first heading
    var where = 0;
    
    //get doc body 
    var body = doc.getBody();

    //big foreach loop
    rows.forEach(function(row, index) {
        
        //with it sorted we have the headers we don't care about just return
        if (index === 0) return;
        
        //first index of data we need to setup dept/where and table headers
        if (index === 1) {
          //add 'Where' as header
            var pageWhereHeader = body.appendParagraph(row[4]+" "+ row[5]).setHeading(DocumentApp.ParagraphHeading.HEADING1);
            
            //center
            pageWhereHeader.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
                
                //add lables to cells
                table = body.appendTable(cells);
                
                //hide cell outline
                table.setBorderColor("#ffffff");

        }
        //we are in the next batch/group of where
        if (where !== row[4]) {
            
            //make sure we are not in the init template
            if (where !== 0) {
                
                //looks like we are switching groups 
                //add page break
                doc.getBody().appendPageBreak()
                
                //add new dept/where
                var pageWhereHeader = body.appendParagraph(row[4]+" "+ row[5]).setHeading(DocumentApp.ParagraphHeading.HEADING1);
                
                //set center
                pageWhereHeader.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
                
                //create new table and add cell headers
                table = doc.getBody().appendTable(cells)
                
                //hide table cell colors
                table.setBorderColor("#ffffff");
            }
            //if we are not in the first groupf we need to init where with what the first dept should be
            where = row[4];
        }
        //for people that are in the same dept
        if (where == row[4]) {
            
            //we need to add 1 row
            for (var i = 0; i < 1; i++) {
                
                //adding the 1 row
                var tr = table.appendTableRow();

                //add 5 cells in each row
                for (var j = 0; j < 5; j++) {
                    
                    //first 4 from data order matters
                    if (j !== 4) {
                        
                        //normal add data from sheet
                        var td = tr.appendTableCell(row[j]);
                    
                    //sig line
                    } else {
                        
                        //looks good enough
                        var td = tr.appendTableCell("____________________________________");
                        
                        //set width of sig line a little longer
                        td.setWidth(250)
                    }
                }
            }
            
            //we should not hit this else but just in case lets log it and set where to the row we are on.
        } else {
            where = row[4]
            Logger.log("else= " + row[4] + " " + where)
        }
    })
    
    // save doc we are done
    doc.saveAndClose()
    
    //get url form doc we just saved
    const url = doc.getUrl()
    
    // clickable pupup _blank opens new tab
    var htmlString = "<base target=\"_blank\">" +
        "<h3>URL</h3><a href=\"" + url + "\">CLICK ME</a>";
    
    //html service 
    var html = HtmlService.createHtmlOutput(htmlString);
    SpreadsheetApp.getUi().showModalDialog(html, 'Open Doc in new tab.');
//end of our function    
}
