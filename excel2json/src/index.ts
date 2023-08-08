import * as fs from 'fs';
import * as ExcelJS from 'exceljs';
//#region  json结构
interface LanguageData{
    languagecode: string;
    ch: string;
    en: string;
}

interface ListOption {
    key: string;
    value?: boolean;
}
  
interface ListItem {
    title: string;
    expression: string;
    type: string;
    group: string;
    suffix:string;
    options: ListOption[];
    tip:ContentItem;
}
interface SubTabData {
    expression: string;
    group: string;
    plus: string;
    list: ListItem[];
}

interface TabData {
    expression: string;
    group: string;
    tabs: Record<string, SubTabData>;
    }

interface MainData {
    [key: string]: string;
}

interface JsonStructure {
    [key: string]: MainData | TabData;
}
interface ContentItem {
    [key: string]: string | string[] | ListOption[] | ContentItem;
  }

type  ContentTable=ContentItem[]
   
interface SingleTableSheetData {
    heads: string[];
    body: ContentTable[];
}
//#endregion

//#region utli


//判断文件路径是否存在
function pathExists(path: string): boolean {
    try {
      fs.accessSync(path);
      return true; // Path exists
    } catch (error) {
      return false; // Path does not exist
    }
  }
//从表达式中提取点名
function extractPointName(expression: string): string | null {
    const pointNamePattern = /[A-Z][A-Za-z0-9_]*/;
    const match = expression.match(pointNamePattern);
    return match ? match[0] : null;
  }
//根据col名称动态获取列
function getCellValue(sheet: ExcelJS.Worksheet, rowNumber: number, columnMap: { [columnKey: string]: number }, columnKey: string): string {
    const columnNumber = columnMap[columnKey];
    if(!columnNumber) return ''
    return sheet.getCell(rowNumber, columnNumber).text;
  }
async function exportJsonToFile(data: any, filePath: string): Promise<void> {    
    fs.writeFile(filePath, JSON.stringify(data, null, 2), 'utf-8',(err)=>{
        console.warn('exportJsonToFile:'+err);
    });
    //console.log(`Exported to ${filePath}` +JSON.stringify(data, null, 2));
  }

function mapToObject(map: Map<string, string>): Record<string, string> {
    return Array.from(map.entries()).reduce((obj, [key, value]) => {
      obj[key] = value;
      return obj;
    },  {} as Record<string, string>);
  }
//#endregion  parse sheet data
//check excel format
async function checkExcelFormat(filePath: string): Promise<boolean> {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);
  
    const requiredSheets = ['languagelist', 'languageconfig', 'tablist'];
  
    for (const sheetName of requiredSheets) {
      const sheet = workbook.getWorksheet(sheetName);
      if (!sheet) {
        console.log(`Sheet "${sheetName}" is missing.`);
        return false;
      }
    }
  
    const languagelistSheet = workbook.getWorksheet('languagelist');
    const languageconfigSheet = workbook.getWorksheet('languageconfig');
    const tablistSheet = workbook.getWorksheet('tablist');
  
    const languagelistHeaders = languagelistSheet.getRow(1).values;
    if (!Array.isArray(languagelistHeaders)||!languagelistHeaders.includes('name') || !languagelistHeaders.includes('address') || !languagelistHeaders.includes('description-ch')) {
      console.log('languagelist sheet headers are incorrect.');
      return false;
    }
  
    const languageconfigHeaders = languageconfigSheet.getRow(1).values;
    if (Array.isArray(languageconfigHeaders) && languageconfigHeaders.includes('default') && languageconfigHeaders.includes('enabled')) {
    // 执行操作
    } else {
    console.log('languageconfig sheet headers are incorrect.');
    return false;
    }
  
    const tablistRowCount = tablistSheet.rowCount;

    for (let rowNumber = 2; rowNumber <= tablistRowCount; rowNumber++) {
    const cell = tablistSheet.getCell(rowNumber, 1);
    if (cell.text === 'list') {
        for (let colNumber = 2; colNumber <= tablistSheet.columnCount; colNumber++) {
        const sheetNameCell = tablistSheet.getCell(rowNumber, colNumber);
        const sheetName = sheetNameCell.text.replace(/"/g, '');
        if(sheetName.trim().length==0) continue;
        const referencedSheet = workbook.getWorksheet(sheetName);
        if (!referencedSheet) {
            console.log(`Sheet "${sheetName}" referenced in tablist is missing.`);
            return false;
        }
        }
    }
    }
  
    return true;
  }
//language config
async function parseLanguageConfig(filePath: string,sheetName:string) : Promise<any[]>{
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);
    const worksheet = workbook.getWorksheet(sheetName);
    if (!worksheet) {
        throw new Error('language_config worksheet not found!');
    }
    const jsonResult: any[] = [];
    const headers: string[] = [];
    // Get headers from the first row, starting from the second column (B)
    for (let colNumber = 2; colNumber <= worksheet.columnCount; colNumber++) {
        const cell = worksheet.getCell(1, colNumber);
        headers.push(cell.text);
    }

    // Process data rows starting from the second row (rowNumber = 2)
    for (let rowNumber = 2; rowNumber <= worksheet.rowCount; rowNumber++) {
        const rowData: any = {};
        const nameCell = worksheet.getCell(rowNumber, 1);
        if (!nameCell.text) {
        continue; // Skip empty rows
        }
        rowData.name = nameCell.text;

        for (let colNumber = 2; colNumber <= worksheet.columnCount; colNumber++) {
        const header = headers[colNumber - 2];
        const cell = worksheet.getCell(rowNumber, colNumber);
        if (cell.text) {
            rowData[header] = cell.text === 'true' || cell.text === 'false' ? cell.text === 'true' : cell.text;
        }
        }

        jsonResult.push(rowData);
    }

    return jsonResult;
}
//languagelist解析并转换 仅支持description-ch  和description-en
async function parseLanguageList(filePath: string,sheetName:string): Promise<{ch: Map<string, string>, en: Map<string, string>}> {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);
    
    const ch = new Map<string, string>();
    const en = new Map<string, string>();
  
    const sheet = workbook.getWorksheet(sheetName);
    if (!sheet) {
      throw new Error('Sheet "languagelist" not found.');
    }
     //languagelist 列名 列号映射
     const columnMap: { [columnKey: string]: number } = {};
     for (let columnNumber = 1; columnNumber <= sheet.columnCount; columnNumber++) {
         const cell = sheet.getCell(1, columnNumber);
         columnMap[cell.text as string] = columnNumber;
     }
    sheet.eachRow((row, rowNumber) => {
      if (rowNumber === 1) return; // Skip header row
      const rowData: LanguageData = {
        languagecode: row.getCell(columnMap['name']).text,
        ch: row.getCell(columnMap['description-ch']).text,
        en: row.getCell(columnMap['description-en']).text,
      };
      if(rowData.ch==""){
        ch.set(rowData.languagecode, rowData.languagecode);
      }
      else{
        ch.set(rowData.languagecode, rowData.ch);
      }
      if(rowData.en==""){
        en.set(rowData.languagecode, rowData.languagecode);
      }
      else{
        en.set(rowData.languagecode, rowData.en);
      }
    });
  
    return { ch, en };
  }
 
//tablist
async function parseTablist(filePath:string,sheetName:string,languageListSheetName:string): Promise<JsonStructure>{
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);
    const jsonStructure: JsonStructure={}
    const worksheet= workbook.getWorksheet(sheetName)
    
    if(!worksheet){
        throw new Error('Tablist worksheet not found!')
    }
    const languageSheet=workbook.getWorksheet(languageListSheetName)
    if(!languageSheet){
        throw new Error('languagelist worksheet not found!')
    }
    //languagelist 列名 列号映射
    const languageListColumnMap: { [columnKey: string]: number } = {};
    for (let columnNumber = 1; columnNumber <= languageSheet.columnCount; columnNumber++) {
        const cell = languageSheet.getCell(1, columnNumber);
        languageListColumnMap[cell.text as string] = columnNumber;
    }
    //process tab_title row
    const tabTitleRow=worksheet.getRow(2)
    tabTitleRow.eachCell((cell,colNumber)=>{
        if(colNumber>1){
            jsonStructure[cell.value as string ]={
                expression:'',
                group:'default',
                tabs:{},
            }
        }
    })
    const columnCount=worksheet.actualColumnCount
    
    for(let colNumber=2;colNumber<=columnCount;colNumber+=1){
        //process other rows   按列循环-每列内再组织tab_title及subtab_title的结构
        for(let rowNumber=4;rowNumber<worksheet.rowCount;rowNumber+=4){
            const subTabTitleCell=worksheet.getCell(rowNumber,colNumber)
            //sub title为空，则舍弃
            if(!subTabTitleCell.value) continue;
            const subTabData: SubTabData ={
                expression:worksheet.getCell(rowNumber+1,colNumber).text,
                group:'default',
                plus:worksheet.getCell(rowNumber+2,colNumber).text,
                list:[],
            }
             //process list data from other sheet
            
            const listSheetName=worksheet.getCell(rowNumber+3,colNumber).text
            //list字段为空的话，就跳过这个list
            if(listSheetName.replace(/"/g, '').trim().length==0){
                console.log("[Tablist]list field is null,but subtitle is exist:"+subTabTitleCell.value)
            }
            else{
                const listSheet=workbook.getWorksheet(listSheetName.replace(/"/g, ''))
                if(!listSheet){
                    throw new Error('List sheet :'+listSheetName+' not found')
                }
                //获取列标识与列位ID的映射
                const listData:ListItem[]=[]
                const columnMap: { [columnKey: string]: number } = {};
                for (let columnNumber = 1; columnNumber <= listSheet.columnCount; columnNumber++) {
                    const cell = listSheet.getCell(1, columnNumber);
                    columnMap[cell.text] = columnNumber;
                }
                    
                for(let rowNumber=2;rowNumber<=listSheet.rowCount;rowNumber++){
                    const title=getCellValue(listSheet, rowNumber, columnMap, 'title');
                    if(!title) continue
                    const expression=getCellValue(listSheet, rowNumber, columnMap, 'expression');
                    const type=getCellValue(listSheet, rowNumber, columnMap, 'type');
                    const group=getCellValue(listSheet, rowNumber, columnMap, 'group');
                    const suffix=getCellValue(listSheet, rowNumber, columnMap, 'suffix');
                //options
                const optionsText=getCellValue(listSheet, rowNumber, columnMap, 'options');
                const optionsMatch=optionsText.match(/\["(.*?)",(\s*.*?)?]/g);
                const options:ListOption[]=[] 
                if(optionsMatch){
                        for(const option of optionsMatch){
                            const [, key, value] = option.match(/\["(.*?)",\s*(.*?)?]/)||[];
                            if(key){
                                options.push({key,value:value?JSON.parse(value):undefined})
                            }
                        }                   
                }
                //tip
                // Process tip data from tag_list sheet
                const tip:ContentItem={} 
                const nameCol=languageListColumnMap['name']
                const addressCol=languageListColumnMap['address']
                const descriptionCol=''
                if(nameCol&&addressCol){
                    const tagName=expression;
                    const tagRow=languageSheet.getColumn(nameCol).values.indexOf(tagName)+1
                    if(tagRow>0){
                        const name=languageSheet.getCell(tagRow,nameCol).text
                        const address=languageSheet.getCell(tagRow,addressCol).text
                        const description='&l('+name+')'
                        tip['name']=name
                        tip['address']=address
                        tip['description']=description
                    }
                }
                listData.push({title,expression,type,group,suffix,options,tip})
                }
                if(subTabData){
                    subTabData.list=listData;
                }
            }
            //           
            const tabTitle=worksheet.getCell(2,subTabTitleCell.col).text
            if(jsonStructure[tabTitle]){
                const tabData=jsonStructure[tabTitle]as TabData
                tabData.tabs[subTabTitleCell.value as string]=subTabData
            }
        }
       
    }
    return jsonStructure
}
//# single table 
async function parseSingleTable(filePath:string,targetJsonDir:string,languageListSheetName:string) :Promise<void>{
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);

    const languageSheet=workbook.getWorksheet(languageListSheetName)
    if(!languageSheet){
        throw new Error('languagelist worksheet not found!')
    }
    //languagelist 列名 列号映射
    const languageListColumnMap: { [columnKey: string]: number } = {};
    for (let columnNumber = 1; columnNumber <= languageSheet.columnCount; columnNumber++) {
        const cell = languageSheet.getCell(1, columnNumber);
        languageListColumnMap[cell.text as string] = columnNumber;
    }

    for (const worksheet of workbook.worksheets){
        if(worksheet.name.startsWith('#')){
            const singleTableSheetData: SingleTableSheetData={
                heads:[],
                body:[],
            }
            //process header row
            const headerRow=worksheet.getRow(2)
            headerRow.eachCell((cell,colNumber)=>{
                if(colNumber>1){
                    const trimmedText = cell.text.trim();
                    const finalText = /^"+$/.test(trimmedText) ? '' : trimmedText;
                    singleTableSheetData.heads.push(finalText)
                }
            })
            //determine the context area based on header type
            const contentAreas:{ startRow: number }[]=[]

            let currentRow=4
            while(true){
                const titleTypeCell=worksheet.getCell(currentRow,1)
                if(titleTypeCell.text=='title'){
                    contentAreas.push({startRow:currentRow})
                }
                currentRow+=1
                if(currentRow>worksheet.rowCount){
                    break;
                }
            }
            //获取列位名称与列位号的映射
            const columnMap: { [columnKey: string]: number } = {};
            for (let columnNumber = 1; columnNumber <= worksheet.columnCount; columnNumber++) {
                const cell = worksheet.getCell(1, columnNumber);
                columnMap[cell.text] = columnNumber;
            }
            // Process content rows  contentAreas 所有内容区域的首行
            for(const contentArea of contentAreas){
                const contentTable: ContentTable=[]
                //遍历所有列 ，每列的从title-type是一个有效内容区域
                for(let colNumber = 2; colNumber <= worksheet.columnCount; colNumber++){
                    let contentItem: ContentItem = {};
                    let contentStartRow=contentArea.startRow
                    let curTitle=''
                    while (true) {
                        //当前行
                        const columnValueTitleCell = worksheet.getCell(contentStartRow, 1);
                        if (!columnValueTitleCell.text ) {
                            break;
                        }
                        if(columnValueTitleCell.text=='title'){

                        }
                        if(columnValueTitleCell.text=='options'){

                        }
                        
                        //根据所在行，进行部分特殊处理
                        switch(columnValueTitleCell.text){
                            case 'title':
                                curTitle= worksheet.getCell(contentStartRow, colNumber).text
                                contentItem[columnValueTitleCell.value as string ] = worksheet.getCell(contentStartRow, colNumber).text;   
                                break;
                            case 'options':
                                //options
                                const optionsText= worksheet.getCell(contentStartRow,colNumber).text;
                                const optionsMatch=optionsText.match(/\["(.*?)",(\s*.*?)?]/g);
                                const options:ListOption[]=[] 
                                if(optionsMatch){
                                        for(const option of optionsMatch){
                                            const [, key, value] = option.match(/\["(.*?)",\s*(.*?)?]/)||[];
                                            if(key){
                                                options.push({key,value:value?JSON.parse(value):undefined})
                                            }
                                        }                   
                                }
                                contentItem[columnValueTitleCell.text as string ] = options;     
                                break;
                            case "type":
                                let typeVal=worksheet.getCell(contentStartRow, colNumber).text;
                                if(typeVal=='title'){
                                    //如果type是title，就只要title,expression,type,和tip  ,默认type是在最后一个，因此直接重新创建contentItem。
                                    contentItem={}
                                    contentItem['title']=curTitle;
                                    contentItem['expression']='';                               
                                    contentItem['type']=worksheet.getCell(contentStartRow, colNumber).text;
                                }
                                else{
                                    contentItem[columnValueTitleCell.value as string ] = worksheet.getCell(contentStartRow, colNumber).text;        
                                }
                                break;
                            default:
                                contentItem[columnValueTitleCell.value as string ] = worksheet.getCell(contentStartRow, colNumber).text;     
                        }
                        contentStartRow++;   
                        //进入到另外一个内容区域的首行了
                        if(worksheet.getCell(contentStartRow, 1).text=='title') break
                    }
                    //tip 
                     // Process tip data from tag_list sheet
                    const tip:ContentItem={} 
                    const nameCol=languageListColumnMap['name']
                    const addressCol=languageListColumnMap['address']
                    const descriptionCol=''
                    if(nameCol&&addressCol){
                        const tagName=extractPointName(contentItem.expression as string);
                        const tagRow=languageSheet.getColumn(nameCol).values.indexOf(tagName)
                        if(tagName==null||tagName.trim().length==0||tagRow<=0){
                            tip['name']=''
                            tip['address']=''
                            tip['description']=''
                        }
                        else{
                            const name=languageSheet.getCell(tagRow,nameCol).text
                            const address=languageSheet.getCell(tagRow,addressCol).text
                            const description='&l('+name+')'
                            tip['name']=name
                            tip['address']=address
                            tip['description']=description
                        }
                        contentItem['tip' as string]=tip
                    }
                    contentTable.push(contentItem);
                }                
                singleTableSheetData.body.push(contentTable);
            }
            // Generate JSON file name based on sheet name
            const jsonFileName = `${worksheet.name.substring(1)}_format.json`;
            // Write sheet data to JSON file
            exportJsonToFile(singleTableSheetData,targetJsonDir+"/"+jsonFileName);
        }
    }
}
export async function transform2json(
    excelPath: string,
    targetJsonDir: string
  ): Promise<void> {
    // const excelPath = './src/Crane(5).xlsx';
    if(!await pathExists(excelPath)){
        console.log('Excel path is not exist!path:'+excelPath)
        return 
    }
    if(!await pathExists(targetJsonDir)){
        console.log('Target json path is not exist!path:'+targetJsonDir)
        return 
    }
    if(!await checkExcelFormat(excelPath)){
        return 
    }
    const { ch, en } = await parseLanguageList(excelPath, 'languagelist');
    const languageListJsonData = { ch: mapToObject(ch), en: mapToObject(en) };
    // const languageListJsonData=await parseLanguageList(excelPath,'languagelist')
    const languageConfigData = await parseLanguageConfig(
      excelPath,
      'languageconfig'
    );
    const tablistJsonData = await parseTablist(
      excelPath,
      'tablist',
      'languagelist'
    );
    //const formattedJsonData = JSON.stringify(jsonData, null, 2);
    exportJsonToFile(languageListJsonData, targetJsonDir + '/languagelist.json');
    exportJsonToFile(languageConfigData, targetJsonDir + '/languageconfig.json');
    exportJsonToFile(tablistJsonData, targetJsonDir + '/tablist.json');
    await parseSingleTable(excelPath,targetJsonDir, 'languagelist');
  }
//#endregion

//ReadExcelAndTransform2Json('./src/Crane.xlsx','./src')
transform2json('./src/Crane(5).xlsx','./src');
