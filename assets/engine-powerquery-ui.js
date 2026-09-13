/* Excel.Flo – Office-Oberflächenbausteine für den Excel-/Power-Query-Nachbau
 *
 * Gemeinsame Helfer für engine-powerquery-excel.js und engine-powerquery-editor.js:
 * DOM-Helfer h(), Menüs (inkl. Untermenüs), Dialoge, Office-Tooltips,
 * Hinweisleiste und ein Menüband-Renderer, der Gruppen bei zu wenig Platz wie
 * in Office schrittweise verkleinert (erst Beschriftungen, dann ganze Gruppe
 * als Dropdown).
 */

(function () {
  "use strict";

  const Icons = window.ExcelFloIcons;

  /* ---------------- DOM ---------------- */

  function h(tag, attrs, children) {
    const node = document.createElement(tag);
    if (attrs) {
      Object.keys(attrs).forEach((key) => {
        const val = attrs[key];
        if (val === undefined || val === null || val === false) return;
        if (key === "class") node.className = val;
        else if (key === "text") node.textContent = val;
        else if (key === "html") node.innerHTML = val;
        else if (key === "on") Object.keys(val).forEach((ev) => node.addEventListener(ev, val[ev]));
        else if (key === "style") Object.assign(node.style, val);
        else if (key === "dataset") Object.assign(node.dataset, val);
        else if (key === "disabled" || key === "checked" || key === "hidden" || key === "selected" || key === "value") node[key] = val;
        else node.setAttribute(key, val === true ? "" : val);
      });
    }
    [].concat(children || []).forEach((c) => {
      if (c === null || c === undefined || c === false) return;
      node.appendChild(typeof c === "string" ? document.createTextNode(c) : c);
    });
    return node;
  }

  function icon(name, size) {
    return h("span", { class: "xl-icon xl-icon--" + (size || 16), html: Icons.get(name, size || 16) });
  }

  function tipAttrs(tip, fallbackTitle) {
    if (!tip && !fallbackTitle) return {};
    return {
      "data-tip-title": (tip && tip.title) || fallbackTitle,
      "data-tip-text": (tip && tip.text) || "",
    };
  }

  /* ---------------- Menüs ---------------- */

  const menuStack = [];

  function closeMenus(fromLevel) {
    const level = fromLevel || 0;
    while (menuStack.length > level) {
      const m = menuStack.pop();
      m.el.remove();
      if (m.onClose) m.onClose();
    }
  }

  function placeMenu(el, rect, prefer) {
    document.body.appendChild(el);
    const mw = el.offsetWidth;
    const mh = el.offsetHeight;
    const vw = document.documentElement.clientWidth;
    const vh = document.documentElement.clientHeight;
    let x;
    let y;
    if (prefer === "right") {
      x = rect.right - 2;
      y = rect.top - 4;
      if (x + mw > vw - 4) x = rect.left - mw + 2;
    } else if (prefer === "point") {
      x = rect.left;
      y = rect.top;
      if (x + mw > vw - 4) x = vw - mw - 4;
    } else {
      x = rect.left;
      y = rect.bottom;
      if (x + mw > vw - 4) x = Math.max(4, rect.right - mw);
    }
    if (y + mh > vh - 4) y = Math.max(4, (prefer === "right" || prefer === "point" ? vh : rect.top) - mh - 4);
    el.style.left = Math.max(4, x) + "px";
    el.style.top = Math.max(4, y) + "px";
  }

  // items: [{ label, icon, disabled, checked, shortcut, submenu: [...], onClick, separator, header }]
  // opts: { rect | anchor, placement: "below" | "right" | "point", level, content: Element, onClose, className }
  function openMenu(opts, items) {
    const level = opts.level || 0;
    closeMenus(level);
    const rect = opts.rect || opts.anchor.getBoundingClientRect();
    const el = h("div", { class: "xl-menu" + (opts.className ? " " + opts.className : ""), role: "menu" });
    el.addEventListener("mousedown", (e) => e.stopPropagation());
    el.addEventListener("contextmenu", (e) => e.preventDefault());

    if (opts.content) {
      el.appendChild(opts.content);
    } else {
      const hasIcons = items.some((it) => it.icon || it.checked !== undefined);
      items.forEach((it) => {
        if (it.separator) {
          el.appendChild(h("div", { class: "xl-menu__sep" }));
          return;
        }
        if (it.header) {
          el.appendChild(h("div", { class: "xl-menu__header", text: it.header }));
          return;
        }
        const row = h(
          "div",
          {
            class: "xl-menu__item" + (it.disabled ? " is-disabled" : "") + (it.submenu ? " has-sub" : ""),
            role: "menuitem",
            "aria-disabled": it.disabled ? "true" : null,
          },
          [
            hasIcons
              ? h("span", { class: "xl-menu__icon" }, [
                  it.checked ? h("span", { class: "xl-menu__check", text: "✓" }) : it.icon ? icon(it.icon, 16) : null,
                ])
              : null,
            h("span", { class: "xl-menu__label", text: it.label }),
            it.shortcut ? h("span", { class: "xl-menu__shortcut", text: it.shortcut }) : null,
            it.submenu ? h("span", { class: "xl-menu__arrow", text: "›" }) : null,
          ]
        );
        if (it.submenu && !it.disabled) {
          const openSub = () => openMenu({ anchor: row, placement: "right", level: level + 1 }, it.submenu);
          row.addEventListener("mouseenter", openSub);
          row.addEventListener("click", openSub);
        } else {
          row.addEventListener("mouseenter", () => closeMenus(level + 1));
          if (!it.disabled) {
            row.addEventListener("click", () => {
              closeMenus(0);
              if (it.onClick) it.onClick();
            });
          }
        }
        el.appendChild(row);
      });
    }

    placeMenu(el, rect, opts.placement);
    menuStack.push({ el, onClose: opts.onClose });
    return el;
  }

  document.addEventListener("mousedown", () => closeMenus(0));
  window.addEventListener("resize", () => closeMenus(0));
  document.addEventListener("keydown", (e) => {
    if (e.key === "Escape" && menuStack.length) closeMenus(0);
  });

  // Scrollen der Seite oder der Tabelle schließt Menüs – Scrollen innerhalb eines
  // Menüs (z. B. Werteliste im Filtermenü) aber nicht.
  window.addEventListener(
    "scroll",
    (e) => {
      if (!menuStack.length) return;
      if (e.target instanceof Node && menuStack.some((m) => m.el.contains(e.target))) return;
      closeMenus(0);
    },
    true
  );

  /* ---------------- Dialoge ---------------- */

  // host: Element mit position:relative (das nachgebaute Fenster).
  // buttons: [{ label, primary, onClick → false verhindert Schließen }]
  function openDialog(host, opts) {
    closeMenus(0);
    const backdrop = h("div", { class: "xl-dialog-backdrop" });
    const errorEl = h("p", { class: "xl-dialog__error", role: "alert" });
    let closed = false;

    function close() {
      if (closed) return;
      closed = true;
      backdrop.remove();
      document.removeEventListener("keydown", onKey, true);
      if (opts.onClose) opts.onClose();
    }

    const buttons = (opts.buttons || []).map((b) =>
      h("button", {
        type: "button",
        class: "xl-btn" + (b.primary ? " xl-btn--primary" : ""),
        text: b.label,
        on: {
          click: () => {
            errorEl.textContent = "";
            const result = b.onClick ? b.onClick({ close, setError: (msg) => (errorEl.textContent = msg) }) : true;
            if (result !== false) close();
          },
        },
      })
    );

    const dialog = h("div", { class: "xl-dialog", role: "dialog", "aria-label": opts.title, style: { width: (opts.width || 480) + "px" } }, [
      h("div", { class: "xl-dialog__titlebar" }, [
        opts.titlebar ? h("span", { class: "xl-dialog__apptitle", text: opts.titlebar }) : h("span"),
        h("button", { type: "button", class: "xl-dialog__close", "aria-label": "Schließen", html: Icons.get("close", 16), on: { click: close } }),
      ]),
      h("div", { class: "xl-dialog__content" }, [
        opts.title ? h("h3", { class: "xl-dialog__title", text: opts.title }) : null,
        opts.subtitle ? h("p", { class: "xl-dialog__subtitle", text: opts.subtitle }) : null,
        opts.body || null,
        errorEl,
      ]),
      buttons.length ? h("div", { class: "xl-dialog__buttons" }, buttons) : null,
    ]);

    function onKey(e) {
      if (e.key === "Escape") {
        e.stopPropagation();
        close();
      } else if (e.key === "Enter" && e.target.tagName !== "TEXTAREA" && e.target.tagName !== "BUTTON") {
        const primary = buttons.find((b) => b.classList.contains("xl-btn--primary"));
        if (primary) {
          e.preventDefault();
          primary.click();
        }
      }
    }

    backdrop.appendChild(dialog);
    backdrop.addEventListener("mousedown", (e) => e.stopPropagation());
    host.appendChild(backdrop);
    document.addEventListener("keydown", onKey, true);
    const first = dialog.querySelector("input:not([type=radio]):not([type=checkbox]), select, textarea") || buttons[0];
    if (first) setTimeout(() => first.focus(), 0);
    return { close, el: dialog };
  }

  // Formular-Helfer für Dialoge
  function field(label, control, extraClass) {
    return h("label", { class: "xl-field" + (extraClass ? " " + extraClass : "") }, [h("span", { class: "xl-field__label", text: label }), control]);
  }

  function select(options, value, onChange) {
    const s = h(
      "select",
      { class: "xl-select", on: { change: () => onChange && onChange(s.value) } },
      options.map((o) => h("option", { value: o.value, text: o.label, disabled: o.disabled }))
    );
    s.value = value;
    return s;
  }

  function input(value, onInput, attrs) {
    const i = h("input", Object.assign({ class: "xl-input", type: "text", autocomplete: "off", spellcheck: "false" }, attrs || {}));
    i.value = value == null ? "" : value;
    if (onInput) i.addEventListener("input", () => onInput(i.value));
    return i;
  }

  function radio(name, label, checked, onChange) {
    const r = h("input", { type: "radio", name, checked });
    r.addEventListener("change", () => r.checked && onChange());
    return h("label", { class: "xl-radio" }, [r, h("span", { text: label })]);
  }

  function checkbox(label, checked, onChange, disabled) {
    const c = h("input", { type: "checkbox", checked, disabled });
    c.addEventListener("change", () => onChange(c.checked));
    return h("label", { class: "xl-checkbox" + (disabled ? " is-disabled" : "") }, [c, h("span", { text: label })]);
  }

  /* ---------------- Tooltips ---------------- */

  let tipTimer = null;
  let tipEl = null;

  function hideTip() {
    clearTimeout(tipTimer);
    if (tipEl) tipEl.remove();
    tipEl = null;
  }

  function bindTooltips(root) {
    root.addEventListener("mouseover", (e) => {
      const target = e.target.closest("[data-tip-title]");
      if (!target || !root.contains(target)) return;
      if (target.contains(e.relatedTarget)) return;
      hideTip();
      tipTimer = setTimeout(() => {
        if (menuStack.length) return;
        const rect = target.getBoundingClientRect();
        tipEl = h("div", { class: "xl-tip", role: "tooltip" }, [
          h("strong", { text: target.dataset.tipTitle }),
          target.dataset.tipText ? h("span", { text: target.dataset.tipText }) : null,
        ]);
        document.body.appendChild(tipEl);
        const vw = document.documentElement.clientWidth;
        tipEl.style.left = Math.min(rect.left, vw - tipEl.offsetWidth - 6) + "px";
        tipEl.style.top = rect.bottom + 4 + "px";
      }, 550);
    });
    root.addEventListener("mouseout", (e) => {
      const target = e.target.closest("[data-tip-title]");
      if (target && !target.contains(e.relatedTarget)) hideTip();
    });
    root.addEventListener("mousedown", hideTip);
  }

  /* ---------------- Hinweisleiste ---------------- */

  function messageBar(message, onClose) {
    if (!message) return null;
    return h("div", { class: "xl-msgbar xl-msgbar--" + (message.kind || "info"), role: "status" }, [
      icon(message.kind === "warning" ? "warning" : "info", 16),
      h("span", { class: "xl-msgbar__title", text: message.title || "HINWEIS" }),
      h("span", { class: "xl-msgbar__text", text: message.text }),
      h("button", { type: "button", class: "xl-msgbar__close", "aria-label": "Hinweis schließen", html: Icons.get("close", 16), on: { click: onClose } }),
    ]);
  }

  /* ---------------- Menüband ---------------- */

  function caret() {
    return h("span", { class: "xl-caret", "aria-hidden": "true" });
  }

  function runControl(ctrl, anchor) {
    if (ctrl.disabled) return;
    if (ctrl.menu && !ctrl.onClick) {
      openMenu({ anchor, placement: "below" }, typeof ctrl.menu === "function" ? ctrl.menu() : ctrl.menu);
    } else if (ctrl.onClick) {
      ctrl.onClick(anchor);
    }
  }

  function bigButton(ctrl) {
    const labelLines = String(ctrl.label).split("\n");
    const isSplit = ctrl.menu && ctrl.onClick;
    const tip = tipAttrs(ctrl.tip, ctrl.label.replace("\n", " "));

    if (isSplit) {
      const wrap = h("div", { class: "xl-rbig xl-rbig--split" + (ctrl.disabled ? " is-disabled" : "") });
      const top = h("button", Object.assign({ type: "button", class: "xl-rbig__top", disabled: ctrl.disabled }, tip), [icon(ctrl.icon, 32)]);
      top.addEventListener("click", () => ctrl.onClick(top));
      const bottom = h("button", { type: "button", class: "xl-rbig__bottom", disabled: ctrl.disabled, "aria-haspopup": "menu" }, [
        h("span", { class: "xl-rbig__label" }, labelLines.map((l, i) => h("span", {}, [l, i === labelLines.length - 1 ? caret() : null]))),
      ]);
      bottom.addEventListener("mousedown", (e) => e.stopPropagation());
      bottom.addEventListener("click", () => openMenu({ anchor: wrap, placement: "below" }, typeof ctrl.menu === "function" ? ctrl.menu() : ctrl.menu));
      wrap.appendChild(top);
      wrap.appendChild(bottom);
      return wrap;
    }

    const btn = h(
      "button",
      Object.assign({ type: "button", class: "xl-rbig" + (ctrl.pressed ? " is-pressed" : ""), disabled: ctrl.disabled, "aria-haspopup": ctrl.menu ? "menu" : null }, tip),
      [
        icon(ctrl.icon, 32),
        h(
          "span",
          { class: "xl-rbig__label" },
          labelLines.map((l, i) => h("span", {}, [l, ctrl.menu && i === labelLines.length - 1 ? caret() : null]))
        ),
      ]
    );
    if (ctrl.menu) btn.addEventListener("mousedown", (e) => e.stopPropagation());
    btn.addEventListener("click", () => runControl(ctrl, btn));
    return btn;
  }

  function smallButton(ctrl, iconOnly) {
    const btn = h(
      "button",
      Object.assign(
        {
          type: "button",
          class: "xl-rsmall" + (iconOnly || ctrl.iconOnly ? " is-icon-only" : "") + (ctrl.pressed ? " is-pressed" : ""),
          disabled: ctrl.disabled,
          "aria-label": ctrl.label,
          "aria-haspopup": ctrl.menu ? "menu" : null,
        },
        tipAttrs(ctrl.tip, ctrl.label)
      ),
      [
        ctrl.icon ? icon(ctrl.icon, 16) : null,
        h("span", { class: "xl-rsmall__label", text: ctrl.label }),
        ctrl.menu ? caret() : null,
      ]
    );
    if (ctrl.menu) btn.addEventListener("mousedown", (e) => e.stopPropagation());
    btn.addEventListener("click", () => runControl(ctrl, btn));
    return btn;
  }

  function renderControl(ctrl) {
    if (ctrl.type === "big") return bigButton(ctrl);
    if (ctrl.type === "stack") {
      return h(
        "div",
        { class: "xl-rstack" + (ctrl.iconOnly ? " is-icon-only" : "") + (ctrl.keepLabels ? " keep-labels" : "") },
        ctrl.items.map((it) => (it.type === "check" ? renderControl(it) : smallButton(it, ctrl.iconOnly)))
      );
    }
    if (ctrl.type === "check") {
      const c = h("input", { type: "checkbox", checked: !!ctrl.checked, disabled: ctrl.disabled });
      c.addEventListener("change", () => ctrl.onChange && ctrl.onChange(c.checked));
      return h("label", Object.assign({ class: "xl-rcheck" + (ctrl.disabled ? " is-disabled" : "") }, tipAttrs(ctrl.tip, ctrl.label)), [c, h("span", { text: ctrl.label })]);
    }
    if (ctrl.type === "gallery") {
      return h("div", { class: "xl-rgallery" }, [
        h(
          "div",
          { class: "xl-rgallery__items" },
          ctrl.items.map((it) =>
            h(
              "button",
              Object.assign({ type: "button", class: "xl-rgallery__item", on: { click: () => it.onClick && it.onClick() } }, tipAttrs(it.tip, it.label)),
              [icon(it.icon, 32), h("span", { text: it.label })]
            )
          )
        ),
        h("div", { class: "xl-rgallery__scroll" }, [
          h("button", { type: "button", class: "xl-rgallery__more", "aria-label": "Weitere", on: { click: () => ctrl.onMore && ctrl.onMore() } }, [caret()]),
        ]),
      ]);
    }
    return smallButton(ctrl);
  }

  function renderGroupFull(group) {
    return h("div", { class: "xl-rgroup__full" }, [
      h("div", { class: "xl-rgroup__controls" }, group.controls.map(renderControl)),
      h("div", { class: "xl-rgroup__label", text: group.label }),
    ]);
  }

  // groups: [{ label, icon (32er für eingeklappte Darstellung), priority (klein = zuerst verkleinert), controls: [...] }]
  function renderRibbon(groups) {
    const row = h("div", { class: "xl-ribbon__row" });
    groups.forEach((group) => {
      const collapsedBtn = h("button", { type: "button", class: "xl-rbig xl-rgroup__collapsed-btn", "aria-haspopup": "menu" }, [
        icon(group.icon || "emptyBig", 32),
        h("span", { class: "xl-rbig__label" }, [h("span", {}, [group.label, caret()])]),
      ]);
      collapsedBtn.addEventListener("mousedown", (e) => e.stopPropagation());
      collapsedBtn.addEventListener("click", () => {
        openMenu({ anchor: collapsedBtn, placement: "below", className: "xl-menu--ribbon" }, null);
        const menu = menuStack[menuStack.length - 1].el;
        menu.appendChild(h("div", { class: "xl-rgroup is-popup" }, [renderGroupFull(group)]));
      });

      const el = h("div", { class: "xl-rgroup", dataset: { priority: group.priority == null ? 50 : group.priority } }, [
        renderGroupFull(group),
        h("div", { class: "xl-rgroup__collapsed" }, [collapsedBtn, h("div", { class: "xl-rgroup__label", text: group.label })]),
      ]);
      row.appendChild(el);
    });
    return row;
  }

  // Verkleinert Gruppen wie Office, bis das Menüband in die Breite passt.
  function fitRibbon(row) {
    if (!row || !row.isConnected) return;
    const groups = Array.from(row.querySelectorAll(":scope > .xl-rgroup"));
    groups.forEach((g) => g.classList.remove("is-compact", "is-collapsed"));
    const order = groups.slice().sort((a, b) => Number(a.dataset.priority) - Number(b.dataset.priority));
    const overflow = () => row.scrollWidth > row.clientWidth + 1;
    for (const g of order) {
      if (!overflow()) return;
      if (g.querySelector(".xl-rstack:not(.is-icon-only)")) g.classList.add("is-compact");
    }
    for (const g of order) {
      if (!overflow()) return;
      g.classList.add("is-collapsed");
    }
  }

  function observeRibbon(row) {
    let lastWidth = 0;
    const ro = new ResizeObserver((entries) => {
      const w = Math.round(entries[0].contentRect.width);
      if (w === lastWidth) return;
      lastWidth = w;
      requestAnimationFrame(() => fitRibbon(row));
    });
    ro.observe(row);
    requestAnimationFrame(() => fitRibbon(row));
    return ro;
  }

  function tabs(names, active, onSelect, extraClassFor) {
    return h(
      "div",
      { class: "xl-tabs", role: "tablist" },
      names.map((name) =>
        h("button", {
          type: "button",
          role: "tab",
          class: "xl-tab" + (name === active ? " is-active" : "") + (extraClassFor ? " " + (extraClassFor(name) || "") : ""),
          "aria-selected": name === active ? "true" : "false",
          text: name,
          on: { click: (e) => onSelect(name, e.currentTarget) },
        })
      )
    );
  }

  window.ExcelFloUI = {
    h,
    icon,
    openMenu,
    closeMenus,
    openDialog,
    field,
    select,
    input,
    radio,
    checkbox,
    bindTooltips,
    hideTip,
    messageBar,
    renderRibbon,
    observeRibbon,
    fitRibbon,
    tabs,
    caret,
  };
})();
