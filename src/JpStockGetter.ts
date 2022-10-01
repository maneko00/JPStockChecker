type Variables = {[key: string]: string | any | undefined}

const kUpdateDateDisplayRow: number = 7;
const kUpdateDateDisplayColumn: number = 2;
const kUrl: string = 'https://kabudata-dll.com/wp-content/uploads/';

export class JpStockGetter {
    public constructor(){}

    /**
     * CSV取得  
     *
     * @return {Array<Array<string>>} CSV
     * @date 2022/9/27
     */
    private GetCsvFile(): Array<Array<string>> 
    {
        let date = new Date();
        let csv: Array<Array<string>> = [];
        let count = 0;
        do{
            let year_month = Utilities.formatDate(date, 'JST', 'yyyy/MM'); //2022/08
            let file = Utilities.formatDate(date, 'JST', 'yyyyMMdd'); //20220822
            const url = kUrl + year_month + '/' + file + '.csv';

            try{
                csv = this.DownloadCsv(url);
            } catch {
                // 前日
                date.setDate(date.getDate()-1);
                count++;
            }
        }while(csv.length == 0 && count < 60)

        return csv;
    }

    /**
     * CSVをダウンロード 
     * 
     * @param {String} url ダウンロードサイトのURL
     * @return {Array<Array<string>>} CSV
     * @date 2022/9/27
     */
    private DownloadCsv(url: string): Array<Array<string>>
    {
        let response = UrlFetchApp.fetch(url);
        let data =  response.getContentText('Shift_JIS');
        return Utilities.parseCsv(data);
    }

    /**
     * 日本株情報取得  
     *
     * @return {void}
     * @date 2022/9/28
     */
    public GetJpStock(): void 
    {
        let csv = this.GetCsvFile();
        if(csv.length == 0)
        {
            Browser.msgBox("情報の更新に失敗しました。");
            return;
        }
        let json_data = this.ConvertCsvToJson(csv);

        const oValues: Array<any> = [];
        json_data.forEach((obj)=> {
            oValues.push([obj.code, obj.name, obj.closing_quotation, obj.dividend, obj.per, obj.pbr]);
        })
        let info_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('株価情報')!;
        info_sheet.getRange(2, 1, json_data.length, 6).setValues(oValues);
    
        // 更新日を入力
        let table_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('一覧')!;
        table_sheet.getRange(kUpdateDateDisplayRow, kUpdateDateDisplayColumn).setValue(json_data[0].current_date + '更新');
        Browser.msgBox("情報を更新しました。");
        return;
    }

    /**
     * CSVをJSONに変換 
     * 
     * @param {Array<Array<string>>} csv CSVデータ
     * @return {Array<Variables>} JSONデータ
     * @date 2022/9/27
     */
    private ConvertCsvToJson(csv: Array<Array<string>>): Array<Variables>
    {
        let titles = csv.shift()!;
        let json_data = csv.map(v => {
            let obj: Variables = {};
            titles.forEach((title, index) => {
                if(title =='市場'){ title = 'market';} 
                if(title =='業種'){ title = 'industry';} 
                if(title =='銘柄コード'){ title = 'code';} 
                if(title =='名称'){ title = 'name';} 
                if(title =='現在日付'){ title = 'current_date';}
                if(title =='前日終値'){ title = 'previous_date_closing_quotation';}
                if(title =='始値'){ title = 'opening_quotation';}
                if(title =='高値'){ title = 'high';}
                if(title =='安値'){ title = 'low';}
                if(title =='終値'){ title = 'closing_quotation';}
                if(title =='出来高'){ title = 'turnover';}
                if(title =='逆日歩'){ title = 'negative_interest_per_diem';}
                if(title =='信用売残'){ title = 'outstanding_sales_on_margin';}
                if(title =='配当'){ title = 'dividend';}
                if(title =='配当落日'){ title = 'ex-dividend_date';}
                if(title =='PER'){ title = 'per';}
                if(title =='PBR'){ title = 'pbr';}

                obj[title] = v[index];
            });
            return obj;
        });
        return json_data;
    }
}