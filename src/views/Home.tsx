import useXLSX from "../hooks/useXLSX"
import { WorkBook, utils, write } from 'xlsx'
import FileSaver from 'file-saver'
import { type } from "os"

function Home() {
    const { handleExcelFileInput } = useXLSX()

    const [source, setSourceFile] = useState<WorkBook>()
    const [comparison, setComparisonFile] = useState<WorkBook>()
    async function handleSourceFile(e: React.ChangeEvent<HTMLInputElement>) {
        const file = await handleExcelFileInput(e)
        setSourceFile(file)
    }
    async function handleComparisonFile(e: React.ChangeEvent<HTMLInputElement>) {
        const file = await handleExcelFileInput(e)
        setComparisonFile(file)
    }
    function getClassNumber(className:string){
        if(className.includes('一')) return 1
        else if(className.includes('二')) return 2
        else if(className.includes('三')) return 3
        else if(className.includes('四')) return 4
        else if(className.includes('五')) return 5
        else if(className.includes('六')) return 6
        else if(className.includes('七')) return 7
        else if(className.includes('八')) return 8
        else if(className.includes('九')) return 9
        else if(className.includes('十')) return 10
    }
    function output() {
        if (!source || !comparison) return alert('請上傳好檔案')
        const sourceSheetsName = source.SheetNames[0]
        const comparisonSheetsName = comparison.SheetNames[0]
        const sourceSheet = source.Sheets[sourceSheetsName]
        const comparisonSheet = comparison.Sheets[comparisonSheetsName]
        let row = 2
        console.log(comparisonSheet)
        while(sourceSheet[`B${row}`]){
            const name =sourceSheet[`B${row}`].v
            const set = sourceSheet[`F${row}`].v
            let compareRow = 2
            while(comparisonSheet[`E${compareRow}`]){
                const compareName = comparisonSheet[`E${compareRow}`].v
                if(name === compareName){
                    const compareClassnumber = getClassNumber(comparisonSheet[`A${compareRow}`].v)
                    const compareSetNumber = comparisonSheet[`B${compareRow}`].v
                    utils.sheet_add_aoa(sourceSheet,[[compareSetNumber]],{origin:`H${row}`})
                    utils.sheet_add_aoa(sourceSheet,[[compareClassnumber]],{origin:`G${row}`})
                    compareRow = 2
                    break
                }
                compareRow++
            }
            row++
        }
        
        const data = write(source,{bookType:'xlsx',type:'file'})
        console.log(typeof(data))
        FileSaver.saveAs(data)
    }
    return (
        <div className='text-white'>
            <header className="text-white">
                <div>
                    <label htmlFor="">來源檔案</label>
                    <input type="file" onChange={handleSourceFile} />
                </div>
                <div>
                    <label htmlFor="">比對檔案</label>
                    <input type="file" onChange={handleComparisonFile} />
                </div>
            </header>
            <button className="border-2 mt-1" onClick={output}>輸出excel</button>
        </div>
    )
}
export default Home