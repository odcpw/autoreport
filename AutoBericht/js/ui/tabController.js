/**
 * Simple tab controller that manages active tab and keyboard navigation.
 * This will evolve to support more complex focus management rules.
 */
export class TabController {
  constructor({ tabs, panels }) {
    this.tabs = Array.from(tabs);
    this.panels = Array.from(panels);
    this.currentIndex = 0;

    this.tabs.forEach((tab, index) => {
      tab.addEventListener('click', () => this.activate(index));
      tab.addEventListener('keydown', (event) => this.handleKeydown(event, index));
    });
  }

  focusFirstTab() {
    const firstTab = this.tabs[0];
    if (firstTab) firstTab.focus();
  }

  activate(index) {
    if (index < 0 || index >= this.tabs.length) return;
    this.currentIndex = index;

    this.tabs.forEach((tab, tabIndex) => {
      const active = tabIndex === index;
      tab.classList.toggle('app-tab--active', active);
      tab.setAttribute('aria-selected', String(active));
      tab.tabIndex = active ? 0 : -1;
    });

    this.panels.forEach((panel, panelIndex) => {
      const active = panelIndex === index;
      panel.classList.toggle('tab-panel--active', active);
      panel.setAttribute('aria-hidden', String(!active));
    });

    const activeTab = this.tabs[index];
    if (activeTab) activeTab.focus();
  }

  handleKeydown(event, index) {
    switch (event.key) {
      case 'ArrowRight':
      case 'ArrowDown':
        event.preventDefault();
        this.moveFocus(index + 1);
        break;
      case 'ArrowLeft':
      case 'ArrowUp':
        event.preventDefault();
        this.moveFocus(index - 1);
        break;
      case 'Home':
        event.preventDefault();
        this.moveFocus(0);
        break;
      case 'End':
        event.preventDefault();
        this.moveFocus(this.tabs.length - 1);
        break;
      case 'Enter':
      case ' ':
        event.preventDefault();
        this.activate(index);
        break;
      default:
        break;
    }
  }

  moveFocus(nextIndex) {
    if (this.tabs.length === 0) return;
    const normalized = (nextIndex + this.tabs.length) % this.tabs.length;
    this.tabs[normalized].focus();
    this.activate(normalized);
  }
}
