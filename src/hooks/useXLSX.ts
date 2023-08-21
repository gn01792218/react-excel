
import { read } from 'xlsx'
export default function useXLSX(){
    

    async function handleExcelFileInput(event:React.ChangeEvent<HTMLInputElement>){
        if(!event.currentTarget.files?.length) return 
        const file = event.currentTarget.files[0]
        const data = await file.arrayBuffer()
        return read(data)
    }
    return {
        handleExcelFileInput,
    }
}