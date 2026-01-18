import ReactDOM from 'react-dom/client';
import '../node_modules/@douyinfe/semi-ui/dist/css/semi.min.css';
import './locales/i18n';
import App from './App';
import LoadApp from './components/LoadApp';

ReactDOM.createRoot(document.getElementById('root') as HTMLElement).render(
  <LoadApp>
    <App />
  </LoadApp>
);
