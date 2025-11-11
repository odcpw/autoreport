const editorRegistry = new WeakMap();

export class MarkdownEditor {
  constructor(textarea, { value = '', onChange } = {}) {
    this.textarea = textarea;
    this.onChange = typeof onChange === 'function' ? onChange : () => {};
    this.cm = null;
    this.handleInput = this.handleInput.bind(this);
    this.textarea.value = value || '';
    this.initEditor();
    editorRegistry.set(this.textarea, this);
  }

  initEditor() {
    const dataset = this.textarea.dataset || {};
    if (typeof window !== 'undefined' && window.CodeMirror && typeof window.CodeMirror.fromTextArea === 'function') {
      this.cm = window.CodeMirror.fromTextArea(this.textarea, {
        mode: 'markdown',
        lineNumbers: false,
        lineWrapping: true,
      });
      this.cm.on('change', () => {
        this.onChange(this.cm.getValue());
      });
      const wrapper = this.cm.getWrapperElement();
      wrapper.dataset.editorWrapper = 'true';
      if (dataset.field) wrapper.dataset.field = dataset.field;
      if (dataset.finding) wrapper.dataset.finding = dataset.finding;
      if (dataset.recommendation) wrapper.dataset.recommendation = dataset.recommendation;
    } else {
      this.textarea.addEventListener('input', this.handleInput);
    }
  }

  handleInput() {
    this.onChange(this.getValue());
  }

  getValue() {
    if (this.cm) {
      return this.cm.getValue();
    }
    return this.textarea.value;
  }

  focus() {
    if (this.cm && typeof this.cm.focus === 'function') {
      this.cm.focus();
    } else {
      this.textarea.focus();
    }
  }

  destroy() {
    if (this.cm && typeof this.cm.toTextArea === 'function') {
      this.cm.toTextArea();
      this.cm = null;
    } else {
      this.textarea.removeEventListener('input', this.handleInput);
    }
    editorRegistry.delete(this.textarea);
  }

  static getInstance(element) {
    if (!element) return null;
    if (editorRegistry.has(element)) {
      return editorRegistry.get(element);
    }
    if (element.dataset?.editorWrapper === 'true') {
      const textarea = element.querySelector('textarea[data-field]');
      if (textarea) {
        return editorRegistry.get(textarea) || null;
      }
    }
    return null;
  }
}
