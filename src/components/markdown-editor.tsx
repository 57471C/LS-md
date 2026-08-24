import {
  forwardRef,
  useEffect,
  useImperativeHandle,
  useRef,
} from "react";
import {
  EditorView,
  keymap,
  lineNumbers,
  highlightActiveLine,
  highlightActiveLineGutter,
  drawSelection,
  dropCursor,
  placeholder as cmPlaceholder,
} from "@codemirror/view";
import { EditorState } from "@codemirror/state";
import {
  defaultKeymap,
  history,
  historyKeymap,
  indentWithTab,
} from "@codemirror/commands";
import { markdown } from "@codemirror/lang-markdown";
import { javascript } from "@codemirror/lang-javascript";
import { json } from "@codemirror/lang-json";
import { html } from "@codemirror/lang-html";
import { css } from "@codemirror/lang-css";
import { python } from "@codemirror/lang-python";
import {
  HighlightStyle,
  syntaxHighlighting,
  bracketMatching,
  type Language,
} from "@codemirror/language";
import { tags } from "@lezer/highlight";
import type { EditorPatch } from "@/lib/markdown/edit";

const jsLang = javascript({ typescript: true }).language;
const jsonLang = json().language;
const htmlLang = html().language;
const cssLang = css().language;
const pythonLang = python().language;

function fencedLanguage(info: string): Language | null {
  const id = info.trim().split(/[\s,:]/, 1)[0]?.toLowerCase() ?? "";
  if (
    id === "js" ||
    id === "javascript" ||
    id === "jsx" ||
    id === "mjs" ||
    id === "cjs" ||
    id === "ts" ||
    id === "typescript" ||
    id === "tsx"
  ) {
    return jsLang;
  }
  if (id === "json") return jsonLang;
  if (id === "html" || id === "htm") return htmlLang;
  if (id === "css") return cssLang;
  if (id === "python" || id === "py") return pythonLang;
  return null;
}

export type MarkdownEditorHandle = {
  getValue: () => string;
  getSelection: () => { start: number; end: number };
  apply: (patch: EditorPatch) => void;
  focus: () => void;
};

type Props = {
  value: string;
  onChange: (value: string) => void;
  onBold: () => void;
  onItalic: () => void;
  onLink: () => void;
  onApply: () => void;
};

const highlight = HighlightStyle.define([
  {
    tag: tags.heading1,
    color: "var(--color-hl-heading)",
    fontWeight: "700",
    fontSize: "1.12em",
  },
  {
    tag: tags.heading2,
    color: "var(--color-hl-heading)",
    fontWeight: "700",
    fontSize: "1.06em",
  },
  { tag: tags.heading3, color: "var(--color-hl-heading)", fontWeight: "600" },
  { tag: tags.heading4, color: "var(--color-hl-heading)", fontWeight: "600" },
  { tag: tags.heading5, color: "var(--color-hl-heading)", fontWeight: "600" },
  { tag: tags.heading6, color: "var(--color-hl-heading)", fontWeight: "600" },
  { tag: tags.heading, color: "var(--color-hl-heading)", fontWeight: "600" },
  { tag: tags.strong, color: "var(--color-hl-heading)", fontWeight: "700" },
  { tag: tags.emphasis, color: "var(--color-hl-em)", fontStyle: "italic" },
  {
    tag: tags.strikethrough,
    color: "var(--color-pane-muted)",
    textDecoration: "line-through",
  },
  { tag: tags.link, color: "var(--color-hl-link)", textDecoration: "underline" },
  { tag: tags.url, color: "var(--color-hl-mark)" },
  {
    tag: tags.monospace,
    color: "var(--color-hl-code)",
    backgroundColor:
      "color-mix(in oklab, var(--color-hl-code) 14%, transparent)",
  },
  { tag: tags.quote, color: "var(--color-pane-muted)", fontStyle: "italic" },
  { tag: tags.processingInstruction, color: "var(--color-hl-mark)" },
  { tag: tags.meta, color: "var(--color-hl-mark)" },
  { tag: tags.contentSeparator, color: "var(--color-hl-mark)" },
  { tag: tags.labelName, color: "var(--color-hl-link)" },
  { tag: tags.comment, color: "var(--color-pane-muted)", fontStyle: "italic" },
  { tag: tags.keyword, color: "var(--color-hl-kw)", fontWeight: "600" },
  { tag: tags.string, color: "var(--color-hl-code)" },
  { tag: tags.number, color: "var(--color-hl-em)" },
  { tag: tags.bool, color: "var(--color-hl-kw)" },
  { tag: tags.literal, color: "var(--color-hl-kw)" },
  { tag: tags.null, color: "var(--color-hl-kw)" },
  { tag: tags.operator, color: "var(--color-pane-muted)" },
  { tag: tags.punctuation, color: "var(--color-pane-muted)" },
  { tag: tags.angleBracket, color: "var(--color-pane-muted)" },
  { tag: tags.typeName, color: "var(--color-hl-kw)" },
  { tag: tags.className, color: "var(--color-hl-em)" },
  { tag: tags.variableName, color: "var(--color-pane-fg)" },
  { tag: tags.propertyName, color: "var(--color-hl-em)" },
  { tag: tags.definition(tags.variableName), color: "var(--color-hl-heading)" },
  {
    tag: tags.function(tags.variableName),
    color: "var(--color-hl-heading)",
  },
  { tag: tags.atom, color: "var(--color-hl-kw)" },
  { tag: tags.tagName, color: "var(--color-hl-kw)" },
  { tag: tags.attributeName, color: "var(--color-hl-em)" },
  { tag: tags.attributeValue, color: "var(--color-hl-code)" },
  { tag: tags.escape, color: "var(--color-hl-em)" },
  { tag: tags.regexp, color: "var(--color-hl-code)" },
  { tag: tags.invalid, color: "var(--color-danger)" },
]);

const editorTheme = EditorView.theme(
  {
    "&": {
      height: "100%",
      backgroundColor: "var(--color-pane)",
      color: "var(--color-pane-fg)",
      fontSize: "0.875rem",
    },
    "&.cm-focused": { outline: "none" },
    ".cm-scroller": {
      overflow: "auto",
      fontFamily: "var(--font-mono)",
      lineHeight: "1.625",
    },
    ".cm-content": {
      caretColor: "var(--color-accent-fg)",
      padding: "12px 16px 24px 8px",
      fontFamily: "var(--font-mono)",
      fontSize: "0.875rem",
      lineHeight: "1.625",
    },
    ".cm-line": { padding: "0 2px" },
    ".cm-gutters": {
      backgroundColor: "var(--color-pane)",
      color: "var(--color-pane-muted)",
      border: "none",
      paddingLeft: "8px",
    },
    ".cm-lineNumbers .cm-gutterElement": {
      minWidth: "2.2rem",
      padding: "0 8px 0 0",
    },
    ".cm-activeLine": {
      backgroundColor: "color-mix(in oklab, var(--color-pane-fg) 4%, transparent)",
    },
    ".cm-activeLineGutter": {
      backgroundColor: "transparent",
      color: "var(--color-pane-fg)",
    },
    ".cm-cursor, .cm-dropCursor": {
      borderLeftColor: "var(--color-accent-fg)",
    },
    "&.cm-focused .cm-selectionBackground, .cm-selectionBackground, .cm-content ::selection":
      {
        background: "color-mix(in oklab, var(--color-accent) 38%, transparent) !important",
      },
    ".cm-placeholder": { color: "var(--color-pane-muted)" },
  },
  { dark: true },
);

export const MarkdownEditor = forwardRef<MarkdownEditorHandle, Props>(
  function MarkdownEditor(
    { value, onChange, onBold, onItalic, onLink, onApply },
    ref,
  ) {
    const parentRef = useRef<HTMLDivElement>(null);
    const viewRef = useRef<EditorView | null>(null);
    const valueRef = useRef(value);
    const onChangeRef = useRef(onChange);
    const commandsRef = useRef({ onBold, onItalic, onLink, onApply });

    valueRef.current = value;
    onChangeRef.current = onChange;
    commandsRef.current = { onBold, onItalic, onLink, onApply };

    useImperativeHandle(ref, () => ({
      getValue: () => viewRef.current?.state.doc.toString() ?? valueRef.current,
      getSelection: () => {
        const view = viewRef.current;
        if (!view) {
          return { start: 0, end: 0 };
        }
        const sel = view.state.selection.main;
        return { start: sel.from, end: sel.to };
      },
      apply: (patch) => {
        const view = viewRef.current;
        if (!view) {
          onChangeRef.current(patch.value);
          return;
        }
        view.dispatch({
          changes: {
            from: 0,
            to: view.state.doc.length,
            insert: patch.value,
          },
          selection: { anchor: patch.start, head: patch.end },
        });
        view.focus();
      },
      focus: () => viewRef.current?.focus(),
    }));

    useEffect(() => {
      const parent = parentRef.current;
      if (!parent) return;

      const view = new EditorView({
        parent,
        state: EditorState.create({
          doc: valueRef.current,
          extensions: [
            lineNumbers(),
            highlightActiveLine(),
            highlightActiveLineGutter(),
            drawSelection(),
            dropCursor(),
            history(),
            bracketMatching(),
            EditorView.lineWrapping,
            cmPlaceholder("# Start with a heading"),
            markdown({ codeLanguages: fencedLanguage }),
            syntaxHighlighting(highlight, { fallback: true }),
            editorTheme,
            EditorView.contentAttributes.of({ "aria-label": "Markdown" }),
            keymap.of([
              {
                key: "Mod-b",
                preventDefault: true,
                run: () => {
                  commandsRef.current.onBold();
                  return true;
                },
              },
              {
                key: "Mod-i",
                preventDefault: true,
                run: () => {
                  commandsRef.current.onItalic();
                  return true;
                },
              },
              {
                key: "Mod-k",
                preventDefault: true,
                run: () => {
                  commandsRef.current.onLink();
                  return true;
                },
              },
              {
                key: "Mod-s",
                preventDefault: true,
                run: () => {
                  commandsRef.current.onApply();
                  return true;
                },
              },
              indentWithTab,
              ...historyKeymap,
              ...defaultKeymap,
            ]),
            EditorView.updateListener.of((update) => {
              if (!update.docChanged) return;
              const next = update.state.doc.toString();
              valueRef.current = next;
              onChangeRef.current(next);
            }),
          ],
        }),
      });
      viewRef.current = view;

      return () => {
        view.destroy();
        viewRef.current = null;
      };
    }, [highlight]);

    useEffect(() => {
      const view = viewRef.current;
      if (!view) return;
      if (view.state.doc.toString() === value) return;
      view.dispatch({
        changes: { from: 0, to: view.state.doc.length, insert: value },
      });
    }, [value]);

    return (
      <div
        ref={parentRef}
        className="cm-host h-full min-h-0 w-full overflow-hidden"
      />
    );
  },
);
