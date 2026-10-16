// A very small DOM stand-in, just enough to run the app's UI code in Node. It does not do layout or CSS.
type Listener = (e: any) => void;

export class FakeNode {
  tagName: string;
  nodeType = 1;
  children: any[] = [];
  parentNode: FakeNode | null = null;
  attrs: Record<string, string> = {};
  listeners: Record<string, Listener[]> = {};
  className = '';
  value = '';
  checked = false;
  disabled = false;
  hidden = false;
  open = false;
  files: any[] = [];
  private _text: string | null = null;
  doc: FakeDocument;
  constructor(tag: string, doc: FakeDocument) { this.tagName = tag.toUpperCase(); this.doc = doc; }
  get classList() {
    const self = this;
    return {
      contains: (c: string) => self.className.split(/\s+/).includes(c),
      add: (c: string) => { if (!self.className.split(/\s+/).includes(c)) self.className = (self.className + ' ' + c).trim(); },
      remove: (c: string) => { self.className = self.className.split(/\s+/).filter((x) => x !== c).join(' '); },
    };
  }
  get textContent(): string {
    if (this._text !== null && this.children.length === 0) return this._text;
    return this.children.map((c) => (c.nodeType === 3 ? c.data : c.textContent)).join('');
  }
  set textContent(v: string) { this.children = []; this._text = String(v); }
  append(...kids: any[]) { for (const k of kids) { k.parentNode?.removeChild?.(k); k.parentNode = this; this.children.push(k); this._text = null; } }
  appendChild(k: any) { this.append(k); return k; }
  removeChild(k: any) { this.children = this.children.filter((c) => c !== k); }
  remove() { this.parentNode?.removeChild(this); this.parentNode = null; }
  replaceChildren(...kids: any[]) { this.children = []; this._text = null; this.append(...kids); }
  setAttribute(k: string, v: string) { this.attrs[k] = v; }
  getAttribute(k: string) { return this.attrs[k] ?? null; }
  removeAttribute(k: string) { delete this.attrs[k]; }
  addEventListener(t: string, fn: Listener) { (this.listeners[t] ||= []).push(fn); }
  fire(type: string, ev: any = {}) {
    const e = { type, target: this, defaultPrevented: false, preventDefault() { this.defaultPrevented = true; }, ...ev };
    let n: FakeNode | null = this;
    while (n) { for (const fn of n.listeners[type] || []) fn(e); n = n.parentNode; }
    for (const fn of this.doc.listeners[type] || []) fn(e); // events bubble up to the document
    return e;
  }
  click() { this.fire('click'); }
  focus() { this.doc.activeElement = this; }
  select() {}
  find(pred: (n: FakeNode) => boolean): FakeNode[] {
    const out: FakeNode[] = [];
    const walk = (n: any) => { if (n.nodeType === 1) { if (pred(n)) out.push(n); n.children.forEach(walk); } };
    walk(this);
    return out;
  }
  byTest(id: string): FakeNode {
    const r = this.find((n) => n.attrs['data-test'] === id);
    if (!r.length) throw new Error('no element with data-test=' + id);
    return r[0];
  }
}

export class FakeDocument {
  listeners: Record<string, Listener[]> = {};
  activeElement: any = null;
  copied: string | null = null;
  createElement(tag: string) { return new FakeNode(tag, this); }
  createTextNode(data: string) { return { nodeType: 3, data, parentNode: null } as any; }
  addEventListener(t: string, fn: Listener) { (this.listeners[t] ||= []).push(fn); }
  execCommand() { return true; }
  fire(type: string, ev: any = {}) {
    const e = { type, defaultPrevented: false, preventDefault() { this.defaultPrevented = true; }, ...ev };
    for (const fn of this.listeners[type] || []) fn(e);
    return e;
  }
}
