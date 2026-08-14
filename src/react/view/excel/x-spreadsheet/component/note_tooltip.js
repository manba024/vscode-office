/* global document, window */
import { h } from './element';
import { cssPrefix } from '../config';

const HIDE_DELAY_MS = 120;
const VIEWPORT_GAP = 8;
const POINTER_GAP = 14;

export default class NoteTooltip {
  constructor() {
    document.querySelectorAll(`.${cssPrefix}-note-tooltip`).forEach(el => el.remove());
    this.hideTimer = null;
    this.contentEl = h('div', `${cssPrefix}-note-tooltip-content`);
    this.el = h('div', `${cssPrefix}-note-tooltip`)
      .attr({ role: 'tooltip', 'aria-hidden': 'true' })
      .child(this.contentEl)
      .hide();
    document.body.appendChild(this.el.el);
    this.el
      .on('mouseenter', () => this.cancelHide())
      .on('mouseleave', () => this.hide());
  }

  show(note, clientX, clientY) {
    this.cancelHide();
    this.contentEl.el.textContent = note.text;
    this.el.attr('aria-hidden', 'false').show();

    const box = this.el.box();
    let left = clientX + POINTER_GAP;
    let top = clientY + POINTER_GAP;
    if (left + box.width > window.innerWidth - VIEWPORT_GAP) {
      left = clientX - box.width - POINTER_GAP;
    }
    if (top + box.height > window.innerHeight - VIEWPORT_GAP) {
      top = window.innerHeight - box.height - VIEWPORT_GAP;
    }
    this.el.offset({
      left: Math.max(VIEWPORT_GAP, left),
      top: Math.max(VIEWPORT_GAP, top),
    });
  }

  cancelHide() {
    if (this.hideTimer !== null) {
      window.clearTimeout(this.hideTimer);
      this.hideTimer = null;
    }
  }

  scheduleHide() {
    this.cancelHide();
    this.hideTimer = window.setTimeout(() => this.hide(), HIDE_DELAY_MS);
  }

  hide() {
    this.cancelHide();
    this.el.attr('aria-hidden', 'true').hide();
  }
}
