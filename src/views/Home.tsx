import useXLSX from "../hooks/useXLSX"
import { WorkBook, utils, write } from 'xlsx'
import FileSaver from 'file-saver'

function Home() {
    const { handleExcelFileInput, getValidExcelTypeStr } = useXLSX()
    const [source, setSourceFile] = useState<WorkBook>()
    const [sourceFileName, setSourceFileName] = useState<string>()
    const [compareFileName, setCompareFileName] = useState<string>()
    const [comparison, setComparisonFile] = useState<WorkBook>()
    async function handleSourceFile(e: React.ChangeEvent<HTMLInputElement>) {
        const file = await handleExcelFileInput(e)
        if (!file) return
        setSourceFile(file)
        setSourceFileName(e.target.files![0].name)
    }
    async function handleComparisonFile(e: React.ChangeEvent<HTMLInputElement>) {
        const file = await handleExcelFileInput(e)
        if (!file) return
        setComparisonFile(file)
        setCompareFileName(e.target.files![0].name)
    }
    function getClassNumber(className: string) {
        if (className.includes('一')) return 1
        else if (className.includes('二')) return 2
        else if (className.includes('三')) return 3
        else if (className.includes('四')) return 4
        else if (className.includes('五')) return 5
        else if (className.includes('六')) return 6
        else if (className.includes('七')) return 7
        else if (className.includes('八')) return 8
        else if (className.includes('九')) return 9
        else if (className.includes('十')) return 10
    }
    function output() {
        if (!source || !comparison) return alert('請上傳好檔案')
        const sourceSheetsName = source.SheetNames[0]
        const comparisonSheetsName = comparison.SheetNames[0]
        const sourceSheet = source.Sheets[sourceSheetsName]
        const comparisonSheet = comparison.Sheets[comparisonSheetsName]
        let row = 2
        try{
            while (sourceSheet[`B${row}`]) {
                const name = sourceSheet[`B${row}`].v
                const set = sourceSheet[`F${row}`].v
                let compareRow = 2
                while (comparisonSheet[`E${compareRow}`]) {
                    const compareName = comparisonSheet[`E${compareRow}`].v
                    if (name === compareName) {
                        const compareClassnumber = getClassNumber(comparisonSheet[`A${compareRow}`].v)
                        const compareSetNumber = comparisonSheet[`B${compareRow}`].v
                        utils.sheet_add_aoa(sourceSheet, [[compareSetNumber]], { origin: `H${row}` })
                        utils.sheet_add_aoa(sourceSheet, [[compareClassnumber]], { origin: `G${row}` })
                        compareRow = 2
                        break
                    }
                    utils.sheet_add_aoa(sourceSheet, [['轉學']], { origin: `G${row}` })
                    compareRow++
                }
                row++
            }
        }catch(e){
            return alert(e+'可能是比對檔案或來源檔案格式錯誤，致使程式無法比對，請確認上傳符合格式內容之excel檔案，詳情請洽國蓓妮')
        }
        const data = write(source, { bookType: 'xlsx', type: 'file' })
        FileSaver.saveAs(data)
    }
    return (
        <div className='text-white'>
            <header className="text-white">
                <p className='text-slate-400 mb-5'>
                    使用須知 : 輸出檔案時，請記得檔案名稱後面要加上 <span className="text-red-300">{getValidExcelTypeStr()}</span> 的副檔名唷~!
                </p>
                <div>
                    <label className="flex cursor-pointer hover:text-slate-500" htmlFor="source">
                        <p className="mr-2">+ 來源檔案</p>
                        <p className="text-slate-500">{sourceFileName}</p>
                    </label>
                    <input className="hidden" id='source' type="file" onChange={handleSourceFile} />
                </div>
                <div >
                    <label className="flex cursor-pointer hover:text-slate-500" htmlFor="compare">
                        <p className="mr-2">+ 比對檔案</p>
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