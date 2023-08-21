function TheHeader() {
    return (
        <div className='text-white'>
           <nav className="flex justify-center">
            <Link className="mr-5" to="/">校護系統與學務系統銜接</Link>
            <Link to="/about">About</Link>
           </nav>
        </div>
    )
}
export default TheHeader