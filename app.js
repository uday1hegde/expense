var XLSX = require('xlsx');


var acctFormula1 = "IF(ISNA(VLOOKUP("
var acctFormuma2 = ',analysis!$A$1:$A$45, 1, FALSE)), "NO", "YES")';

var infile = process.argv[2];
var outfile = process.argv[3];
var workbook = XLSX.readFile(infile, {cellDates:true});

var assignmentsSheetName = 'assignments';
var bankSheetName = 'bank';
var creditSheetName = 'credit';

var assignments = getAssignments(workbook.Sheets[assignmentsSheetName]);

var notFoundStrings = [];

console.log("handling bank");
processWorkSheet(workbook.Sheets[bankSheetName], assignments);

console.log("handling credit");
processWorkSheet(workbook.Sheets[creditSheetName], assignments);

addNotFoundStrings(workbook.Sheets[assignmentsSheetName]);

XLSX.writeFile(workbook, outfile, {cellDates:true,bookType:"xlsx"});




function getAssignments(workSheet) {
    var hdrDesc = 'description';
    var hdrCat = 'category';
    var hdrSubCat = 'subcategory';  
    var colAlpha=['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M'];
    var colHdrs={};
    var assignments = new Array();
    var category;
    var subCategory;
    
    //Row 1.
    
    for (var cols=0; cols<colAlpha.length; cols++) {
        if (workSheet[colAlpha[cols]+1] !== undefined) {
            colHdrs[workSheet[colAlpha[cols]+1].v] = colAlpha[cols];
        }
    }
    
    for (var rows = 2; rows < 10000; rows++) { 
        
        if (workSheet[colHdrs[hdrDesc]+rows] === undefined) {
            break;
        }
        description  = workSheet[colHdrs[hdrDesc]+rows].v;
        if (workSheet[colHdrs[hdrCat]+rows] !== undefined) {
            category = workSheet[colHdrs[hdrCat]+rows].v;
        }
        else {
            category = undefined;
        }
        
        if (workSheet[colHdrs[hdrSubCat]+rows] !== undefined) {
            subCategory = workSheet[colHdrs[hdrSubCat]+rows].v;
        }
        else {
            subCategory = undefined;
        }
        
        assignments[description] = [category, subCategory];
    }
    console.log(`found ${rows} assignments`);
    return assignments;
}

function addNotFoundStrings(workSheet) {
   var hdrDesc = 'description';
    var hdrCat = 'category';
    var hdrSubCat = 'subcategory';  
    var colAlpha=['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M'];
    var colHdrs={};
   
    
    //Row 1.
    
    for (var cols=0; cols<colAlpha.length; cols++) {
        if (workSheet[colAlpha[cols]+1] !== undefined) {
            colHdrs[workSheet[colAlpha[cols]+1].v] = colAlpha[cols];
            lastColumn = colAlpha[cols];
        }
    }
    
    for (var rows = 2; rows < 10000; rows++) { 
  
        if (workSheet[colHdrs[hdrDesc]+rows] === undefined) {
            break;
        }
    }

    for (key in notFoundStrings) {
        workSheet[colHdrs[hdrDesc]+rows] = {t:'s', v:key};
        rows++;
    }
    var endOfSheet = lastColumn+rows;
    workSheet['!ref'] = 'A1:'+endOfSheet;
    console.log(`current ref ${workSheet['!ref']}`);
}

function processWorkSheet(workSheet, assignments) {
    
    var notfound = 0;
    var totFound = 0;
    var colAlpha=['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M'];
    var colHdrs={};

    var hdrDate='date';
    var hdrAmount = 'amount';
    var hdrDesc = 'description';
    var hdrCat = 'category';
    var hdrSubCat = 'subcategory';
    var hdrAccounted = 'accounted';
    //Row 1.
    
    for (var cols=0; cols<colAlpha.length; cols++) {
        if (workSheet[colAlpha[cols]+1] !== undefined) {
            colHdrs[workSheet[colAlpha[cols]+1].v] = colAlpha[cols];
        }
    }

    for (var rows = 2; rows < 10000; rows++) { 
        
        var found = false;
                
        if (workSheet[colHdrs[hdrDate]+rows] === undefined) {
            break;
        }

        if (workSheet[colHdrs[hdrAmount]+rows] !== undefined) {
            expense = workSheet[colHdrs[hdrDesc]+rows].v;
            expense = expense.toLowerCase();

            for (var key in assignments) {
                if (expense.includes(key)) {
                    thisCategory = assignments[key][0];
                    thisSub = assignments[key][1];
                    
                    found=true;
                    totFound++;
                    break;
                }
            }
            
    
            if (found == true) {
                if (workSheet[colHdrs[hdrCat]+rows] !== undefined) {
                    currentCategory = workSheet[colHdrs[hdrCat]+rows].v;
                    if (currentCategory != thisCategory) {
                        console.log(`row ${rows} make manual change from category ${currentCategory} to ${thisCategory}`);
                    }
                    if (workSheet[colHdrs[hdrSubCat]+rows] !== undefined) {
                        var currentSubCat =workSheet[colHdrs[hdrSubCat]+rows].v;
                        if (currentSubCat != thisSub) {
                            console.log(`row ${rows} make manual change from subcategory ${currentSubCat} to ${thisSub}`);
                        }   
                    }
                }
                else {                
                    workSheet[colHdrs[hdrCat]+rows] = {t:'s', v:thisCategory};
                    workSheet[colHdrs[hdrSubCat]+rows] = {t:'s', v:thisSub};                    
                }

            } 
            else {
                
                if (workSheet[colHdrs[hdrCat]+rows] === undefined) {
                    console.log(`not found ${expense}`);
                    for (var key in assignments) {
                        if (expense.includes(key)) {
                            console.log(`found ${expense} in ${key}`);
                            break;
                        }
                        else {
                            console.log(`key ${key} not in ${expense}`);
                        }
                    }
                    notfound++;
                    if (notFoundStrings[expense] === undefined) {
                        notFoundStrings[expense] = 1;
                    }
                }
            }
                        
            if(workSheet[colHdrs[hdrAccounted]+rows] === undefined) {
                var formula = acctFormula1+colHdrs[hdrCat]+rows+acctFormuma2;
                workSheet[colHdrs[hdrAccounted]+rows] = {t:'f', f:formula};
            }
        }
    }
    console.log(`found ${totFound} not found ${notfound}`);
}

