// Geist (OFL-1.1) for the shell, self-hosted: the demo makes no font request to a third party.
import '@fontsource-variable/geist';
import './styles.css';
import { createApp, h } from 'vue';
import App from './App.vue';
import PreviewBanner from '../../shared/PreviewBanner.vue';

createApp({
  render: () => h('div', { class: 'demo-frame' }, [h(PreviewBanner), h(App)]),
}).mount('#app');
