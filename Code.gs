    /*


LOOKAHEAD ////////////////////////////////
(relative to position of function itself)
(positive or negative compatible)
(requires iterative calculation to be enabled in settings)
(I used this to duplicate column sections, simulating a nested arrayformula inside an arrayformula)

'=IF(CONCAT(T(CHAR((COLUMN()+8+64))),INDEX(MAP(C:C, LAMBDA(JAW, ROW(JAW))),ROW()+9,1)) = "K13", "FOUND","UGG")'

NESTED ARRAY/////
(this can make distinct arrays extending from multiple headers in a column)

=ARRAY_CONSTRAIN(vstack(TRANSPOSE(SPLIT(left(regexreplace(JOIN("joiner",F2:F),".","1∎"), sum((match(F1,F2:F,0)*2),-2)),"∎"))), match(F1,F2:F,0)-1,1)

COMBINATION OF ABOVE


=if(
  column is first, 
  choose(
    len((row()-1)/127), , indirect(CONCAT(T(CHAR((COLUMN()+64))),INDEX(MAP(A:A, LAMBDA(JAW, ROW(JAW)+1)),1))), indirect(CONCAT(T(CHAR((COLUMN()+64))),INDEX(MAP(C:C, LAMBDA(JAW, ROW(JAW))),ROW()-N(127*((row()-1)/127)),1))))

SIMPLE STRING OCCURENCE COUNTER////////////////////////
(This should be revised using the choose function)

=IF(OR(
NE(DIVIDE(Len(H18) - len(SUBSTITUTE(H18, I17,CHAR(8733))),11),1) ,
NE(DIVIDE(Len(H18) - len(SUBSTITUTE(H18, PROPER(I17),CHAR(8733))),11), 1),
NE(DIVIDE(Len(H18) - len(SUBSTITUTE(H18, UPPER(I17),CHAR(8733))),11),1)),
"MORE THAN OR FEWER THAN 1 OCCURENCE OF "&CELL("CONTENTS", I17)&"; MUST REVIEW",
IF(
DIVIDE(Len(H18) - len(SUBSTITUTE(H18, I17,CHAR(8733))),11) = 1,
SPLIT(SUBSTITUTE(H18, I17,CHAR(8733)), CHAR(8733),FALSE),
IF(
DIVIDE(Len(H18) - len(SUBSTITUTE(H18, PROPER(I17),CHAR(8733))),11) = 1,
SPLIT(SUBSTITUTE(H18, PROPER(I17),CHAR(8733)), CHAR(8733),FALSE),
IF(
DIVIDE(Len(H18) - len(SUBSTITUTE(H18, UPPER(I17),CHAR(8733))),11) = 1,
SPLIT(SUBSTITUTE(H18, UPPER(I17),CHAR(8733)), CHAR(8733),FALSE)))))

ALPHABETIC LIST////

"A","B","C","D","E","F","G","H","I","J","K","L","M","N","O","P","Q","R","S","T","U","V","W","X","Y","Z","AA","AB","AC","AD","AE","AF","AG","AH","AI","AJ","AK","AL","AM","AN","AO","AP","AQ","AR","AS","AT","AU","AV","AW","AX","AY","AZ","BA","BB","BC","BD","BE","BF","BG","BH","BI","BJ","BK","BL","BM","BN","BO","BP","BQ","BR","BS","BT","BU","BV","BW","BX","BY","BZ","CA","CB","CC","CD","CE","CF","CG","CH","CI","CJ","CK","CL","CM","CN","CO","CP","CQ","CR","CS","CT","CU","CV","CW","CX","CY","CZ","DA","DB","DC","DD","DE","DF","DG","DH","DI","DJ","DK","DL","DM","DN","DO","DP","DQ","DR","DS","DT","DU","DV","DW","DX","DY","DZ","EA","EB","EC","ED","EE","EF","EG","EH","EI","EJ","EK","EL","EM","EN","EO","EP","EQ","ER","ES","ET","EU","EV","EW","EX","EY","EZ","FA","FB","FC","FD","FE","FF","FG","FH","FI","FJ","FK","FL","FM","FN","FO","FP","FQ","FR","FS","FT","FU","FV","FW","FX","FY","FZ","GA","GB","GC","GD","GE","GF","GG","GH","GI","GJ","GK","GL","GM","GN","GO","GP","GQ","GR","GS","GT","GU","GV","GW","GX","GY","GZ","HA","HB","HC","HD","HE","HF","HG","HH","HI","HJ","HK","HL","HM","HN","HO","HP","HQ","HR","HS","HT","HU","HV","HW","HX","HY","HZ","IA","IB","IC","ID","IE","IF","IG","IH","II","IJ","IK","IL","IM","IN","IO","IP","IQ","IR","IS","IT","IU","IV","IW","IX","IY","IZ","JA","JB","JC","JD","JE","JF","JG","JH","JI","JJ","JK","JL","JM","JN","JO","JP","JQ","JR","JS","JT","JU","JV","JW","JX","JY","JZ","KA","KB","KC","KD","KE","KF","KG","KH","KI","KJ","KK","KL","KM","KN","KO","KP","KQ","KR","KS","KT","KU","KV","KW","KX","KY","KZ","LA","LB","LC","LD","LE","LF","LG","LH","LI","LJ","LK","LL","LM","LN","LO","LP","LQ","LR","LS","LT","LU","LV","LW","LX","LY","LZ","MA","MB","MC","MD","ME","MF","MG","MH","MI","MJ","MK","ML","MM","MN","MO","MP","MQ","MR","MS","MT","MU","MV","MW","MX","MY","MZ","NA","NB

*/
const SpreadsheetObj = SpreadsheetApp.getActiveSpreadsheet();
/////////
function changeSheetName() {
  const sheet = SpreadsheetObj.getSheetByName('Loop');
  sheet.setName('VLoop')};
///////////
function SheetDelete() {
                            Logger.log (SpreadsheetObj.getSheetName());
                            Logger.log (SpreadsheetObj.getActiveSheet().getName);
  var SheetsNameObjArray = SpreadsheetObj.getSheets();
  var refocusTo = SpreadsheetObj.setActiveSheet(SheetsNameObjArray[5]);
                            Logger.log(refocusTo.getName);
  var Execution = SpreadsheetObj.deleteActiveSheet();
                            Logger.log(Execution.getName)};
////////////
function SheetReorder() {
var SheetsNameObjArray = SpreadsheetObj.getSheets();
SpreadsheetObj.setActiveSheet(SheetsNameObjArray[3]);
SpreadsheetObj.moveActiveSheet(0)};
////////////////
function ChangeFormulasToSingle() {
SpreadsheetObj.getRange("VWrap!A1:L58").activate();
SpreadsheetApp.getActiveRange().setFormula(`
=INDIRECT("'Edit'!"&CHAR(COLUMN()+1-((ROUNDUP(COLUMN() * 0.25)-1)*4)+64)&ROW()+(ROUNDUP(COLUMN() * 0.25)-1)*58&regexextract(T("
  [VAR COLUMNOFFSET = 1, VAR ROWOFFSET = 0, VAR COLUMNS = 4, VAR ROWS =58]: 
  =indirect(""'edit'!""&char(column()+(COLUMNOFFSET)-((roundup(column() * (1/COLUMNS))-1)*COLUMNS)+64)&row()+(roundup(column() * (1/COLUMNS))-1)*ROWS)"),""))
`)};
//note that regex has not been made compatible for column indexes greater than 26
////////////
function ResizeCols() {
var sheets = SpreadsheetObj.getSheets()[0];
//if sheets 
SpreadsheetObj.setColumnWidths()};
















