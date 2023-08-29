import Routers from './router/Routers'
import TheHeader from './components/TheHeader'
function App() {
  return (
    <div className="App bg-black h-screen">
      <TheHeader/>
      <Routers />
    </div>
  )
}

export default App
