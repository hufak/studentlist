import { createApp } from 'vue';
import './tokens.css';
import { setupEmbedding } from './embed';
import App from './App.vue';

setupEmbedding();
createApp(App).mount('#root');
