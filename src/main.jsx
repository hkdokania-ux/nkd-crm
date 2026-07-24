import { StrictMode, Component } from 'react'
import { createRoot } from 'react-dom/client'
import './index.css'
import App from './App.jsx'

class ErrorBoundary extends Component {
  constructor(props){super(props);this.state={err:null};}
  static getDerivedStateFromError(e){return{err:e};}
  render(){
    if(this.state.err){
      return(
        <div style={{padding:30,fontFamily:'monospace',background:'#fff0f0',minHeight:'100vh'}}>
          <h2 style={{color:'#ef4444'}}>⚠️ App crashed</h2>
          <pre style={{whiteSpace:'pre-wrap',fontSize:13,color:'#1e293b'}}>{String(this.state.err)}</pre>
          <pre style={{whiteSpace:'pre-wrap',fontSize:11,color:'#64748b'}}>{this.state.err?.stack}</pre>
          <button onClick={()=>{localStorage.clear();window.location.reload();}} style={{marginTop:20,padding:'10px 20px',background:'#ef4444',color:'#fff',border:'none',borderRadius:8,cursor:'pointer',fontSize:14}}>🗑️ Clear Cache & Reload</button>
        </div>
      );
    }
    return this.props.children;
  }
}

createRoot(document.getElementById('root')).render(
  <StrictMode>
    <ErrorBoundary>
      <App />
    </ErrorBoundary>
  </StrictMode>,
)
