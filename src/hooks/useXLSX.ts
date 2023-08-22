
import { read } from 'xlsx'
export default function useXLSX(){
    async function handleExcelFileInput(event:React.ChangeEvent<HTMLInputElement>){
        if(!event.currentTarget.files?.length) return 
        const fileName = event.currentTarget.files[0].name
        if(!fileName.includes("xlsx") && !fileName.includes('xls')) {
            return alert('請上傳excel使用的.xlsx或 .xls檔案')
        }
        const file = event.currentTarget.files[0]
        const data = await file.arrayBuffer()
        return read(data)
    }
    return {
        handleExcelFileInput,
    }
}