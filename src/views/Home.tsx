import useXLSX from "../hooks/useXLSX"
import { WorkBook, utils, writeFileXLSX } from 'xlsx'

function Home() {
    const { handleExcelFileInput, getValidExcelTypeStr } = useXLSX()
    const [source, setSourceFile] = useState<WorkBook>()
    const [sourceFileName, setSourceFileName] = useState<string>()
    const [compareFileName, setCompareFileName] = useState<string>()
    const [comparison, setComparisonFile] = useState<WorkBook>()
    async function output() {
        if (!source || !comparison) return alert('請上傳好檔案')
        const sourceSheetsName = source.SheetNames[0]
        const comparisonSheetsName = comparison.SheetNames[0]
        const sourceSheet = source.Sheets[sourceSheetsName]
        const comparisonSheet = comparison.Sheets[comparisonSheetsName]
        let row = 2
        try{
            while (sourceSheet[`A${row}`] || sourceSheet[`B${row}`]) {
                let compareRow = 2
                while (comparisonSheet[`E${compareRow}`]) {
                    if(!sourceSheet[`A${row}`]) {
                        utils.sheet_add_aoa(sourceSheet, [['查無資料']], { origin: `G${row}` })
                        compareRow++
                        break
                    }
                    const studentNumber =sourceSheet[`A${row}`].v
                    const compareStudentNumber = comparisonSheet[`C${compareRow}`].v
                    if (studentNumber === compareStudentNumber) {
                        const compareClassnumber = getClassNumber(comparisonSheet[`A${compareRow}`].v)
                        const compareSetNumber = comparisonSheet[`B${compareRow}`].v
                        utils.sheet_add_aoa(sourceSheet, [[compareSetNumber]], { origin: `H${row}` })
                        utils.sheet_add_aoa(sourceSheet, [[compareClassnumber]], { origin: `G${row}` })
                        compareRow = 2
                        break
                    }
                    utils.sheet_add_aoa(sourceSheet, [['查無資料']], { origin: `G${row}` })
                    compareRow++
                }
                row++
            }
        }
        catch(e){
            return alert(e+`
            來源檔案發生錯誤的行數:${row}，
            可能是比對檔案或來源檔案格式錯誤，或資料有誤，
            致使程式無法比對，
            請確認上傳符合格式內容之excel檔案。`,)
        }
        writeFileXLSX(source!,'nurse_system.xls')  //下載檔案
    }
    async function handleSourceFile(e: React.ChangeEvent<HTMLInputElement>) {
        const file = await handleExcelFileInput(e)
        if (!validSourceFile(file!)) return alert('請上傳正確的SSHIS檔案格式:A欄title只能是學號，或學生學號')
        setSourceFile(file)
        setSourceFileName(e.target.files![0].name)
    }
    async function handleComparisonFile(e: React.ChangeEvent<HTMLInputElement>) {
        const file = await handleExcelFileInput(e)
        if (!validCompareFile(file!)) return alert('請上傳正確的學務系統檔案格式，確認C欄title為學號')
        setComparisonFile(file)
        setCompareFileName(e.target.files![0].name)
    }
    function getClassNumber(className: string) {
        const classStr = className.substring(2)
        if (classStr.includes('一')) return 1
        else if (classStr.includes('二')) return 2
        else if (classStr.includes('三')) return 3
        else if (classStr.includes('四')) return 4
        else if (classStr.includes('五')) return 5
        else if (classStr.includes('六')) return 6
        else if (classStr.includes('七')) return 7
        else if (classStr.includes('八')) return 8
        else if (classStr.includes('九')) return 9
        else if (classStr.includes('十')) return 10
    }
    function validSourceFile(file:WorkBook){
        if(!file) return false
        const sheet = file.Sheets[file.SheetNames[0]] 
        const colA = sheet["A1"].v
        if( colA!== "學號" && colA!=="學生學號" ) return false
        return true
    }
    function validCompareFile(file:WorkBook){
        if(!file) return false
        const sheet = file.Sheets[file.SheetNames[0]] 
        const colC = sheet["C1"].v
        if( colC!== "學號" && colC!=="學生學號" ) return false
        return true
    }
    return (
        <div className='text-white text-left flex flex-col items-center p-5'>
            <header className="text-white">
                <p className='text-slate-400 mb-5'>
                    使用須知 :請在SSHIS舊班excel檔案中的A欄增加 <span className="text-red-300">學號</span>欄位，才能順利比對
                </p>
                <div>
                    <label className="flex items-center cursor-pointer hover:text-slate-500 mb-2" htmlFor="source">
                        <p className="mr-2 p-2 border-2">+ SSHIS-舊班檔案</p>
                        <p className="text-slate-500">{sourceFileName}</p>
                    </label>
                    <input className="hidden" id='source' type="file" onChange={handleSourceFile} />
                </div>
                <div >
                    <label className="flex items-center cursor-pointer hover:text-slate-500" htmlFor="compare">
                        <p className="mr-2 p-2 border-2">+ 學務系統-新班檔案</p>
                        <p className="text-slate-500">{compareFileName}</p>
                    </label>
                    <input className="hidden" id='compare' type="file" onChange={handleComparisonFile} />
                </div>
            </header>
            <button className="w-[150px] h-[50px] border-2 my-5 ml-10 hover:bg-slate-600" onClick={output}>輸出excel</button>
        </div>
    )
}
export default Home