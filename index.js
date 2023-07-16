const XLSX = require('xlsx')
const GLOB = require("glob");

main()

async function main(){
  // inputフォルダ
  const infilesPath = "infiles"
  // outputフォルダ
  const outfilesPath = "outfiles"
  // outputファイル名
  const outfile = "output.xlsx"
  
  const outputData = []
  const outputHeader = ["fileName", "Key", "Value"]
  outputData.push(outputHeader)

  // inputファイル一覧
  const files = GLOB.sync(infilesPath + "/*.xlsx")

  files.sort()

  // inputファイル読み込み
  files.forEach(file => {
    //  エクセル読込
    const workbook = XLSX.readFile(file)
    const sheet = workbook.Sheets["data"]
    const rows = XLSX.utils.sheet_to_json(sheet)
    
    //  出力データ作成
    rows.forEach(row => {
      const data = [file, row.Key, row.Value]
      outputData.push(data)
    })
  })

  // エクセルファイルデータ作成
  const outBook = XLSX.utils.book_new()
  const outSheet = XLSX.utils.aoa_to_sheet(outputData)
  const outSheetName = "output"

  // ファイル出力
  XLSX.utils.book_append_sheet(outBook, outSheet, outSheetName)
  XLSX.writeFile(outBook, outfilesPath + "/" + outfile)

}