import { useEffect, useRef } from 'react';
import { Button, makeStyles, tokens, Tooltip } from '@fluentui/react-components';
import {
  TextBold20Regular,
  TextItalic20Regular,
  TextUnderline20Regular,
  TextHeader120Regular,
  TextHeader220Regular,
  TextBulletListLtr20Regular,
  TextNumberListLtr20Regular,
  Link20Regular,
} from '@fluentui/react-icons';
import { sanitizeArticleHtml } from '../processing/pipeline';

const useStyles = makeStyles({
  root: {
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalS,
  },
  toolbar: {
    display: 'flex',
    flexWrap: 'wrap',
    gap: tokens.spacingHorizontalXS,
    padding: tokens.spacingHorizontalS,
    borderTopWidth: '1px',
    borderRightWidth: '1px',
    borderBottomWidth: '1px',
    borderLeftWidth: '1px',
    borderTopStyle: 'solid',
    borderRightStyle: 'solid',
    borderBottomStyle: 'solid',
    borderLeftStyle: 'solid',
    borderTopColor: tokens.colorNeutralStroke2,
    borderRightColor: tokens.colorNeutralStroke2,
    borderBottomColor: tokens.colorNeutralStroke2,
    borderLeftColor: tokens.colorNeutralStroke2,
    borderRadius: tokens.borderRadiusLarge,
    backgroundColor: tokens.colorNeutralBackground2,
  },
  editor: {
    minHeight: '320px',
    padding: tokens.spacingHorizontalL,
    borderTopWidth: '1px',
    borderRightWidth: '1px',
    borderBottomWidth: '1px',
    borderLeftWidth: '1px',
    borderTopStyle: 'solid',
    borderRightStyle: 'solid',
    borderBottomStyle: 'solid',
    borderLeftStyle: 'solid',
    borderTopColor: tokens.colorNeutralStroke2,
    borderRightColor: tokens.colorNeutralStroke2,
    borderBottomColor: tokens.colorNeutralStroke2,
    borderLeftColor: tokens.colorNeutralStroke2,
    borderRadius: tokens.borderRadiusLarge,
    backgroundColor: tokens.colorNeutralBackground1,
    boxShadow: tokens.shadow2,
    outlineStyle: 'none',
    overflow: 'auto',
    lineHeight: tokens.lineHeightBase400,
  },
});

interface ToolbarAction {
  label: string;
  icon: JSX.Element;
  command: () => void;
}

export interface RichTextEditorProps {
  value: string;
  onCommit: (html: string) => void;
}

export function RichTextEditor({ value, onCommit }: RichTextEditorProps) {
  const s = useStyles();
  const editorRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    if (editorRef.current && editorRef.current.innerHTML !== value) {
      editorRef.current.innerHTML = value;
    }
  }, [value]);

  function focusEditor() {
    editorRef.current?.focus();
  }

  function runCommand(command: string, valueToApply?: string) {
    focusEditor();
    document.execCommand(command, false, valueToApply);
  }

  function createLink() {
    const url = window.prompt('Enter the link URL', 'https://');
    if (!url) {
      return;
    }
    runCommand('createLink', url);
    commit();
  }

  function commit() {
    const cleaned = sanitizeArticleHtml(editorRef.current?.innerHTML ?? value);
    if (editorRef.current && editorRef.current.innerHTML !== cleaned) {
      editorRef.current.innerHTML = cleaned;
    }
    onCommit(cleaned);
  }

  const actions: ToolbarAction[] = [
    { label: 'Bold', icon: <TextBold20Regular />, command: () => runCommand('bold') },
    { label: 'Italic', icon: <TextItalic20Regular />, command: () => runCommand('italic') },
    { label: 'Underline', icon: <TextUnderline20Regular />, command: () => runCommand('underline') },
    { label: 'Heading 1', icon: <TextHeader120Regular />, command: () => runCommand('formatBlock', '<h1>') },
    { label: 'Heading 2', icon: <TextHeader220Regular />, command: () => runCommand('formatBlock', '<h2>') },
    { label: 'Bulleted list', icon: <TextBulletListLtr20Regular />, command: () => runCommand('insertUnorderedList') },
    { label: 'Numbered list', icon: <TextNumberListLtr20Regular />, command: () => runCommand('insertOrderedList') },
    { label: 'Link', icon: <Link20Regular />, command: createLink },
  ];

  return (
    <div className={s.root}>
      <div className={s.toolbar}>
        {actions.map(action => (
          <Tooltip key={action.label} content={action.label} relationship="label">
            <Button appearance="subtle" icon={action.icon} onClick={action.command} />
          </Tooltip>
        ))}
      </div>
      <div
        ref={editorRef}
        className={s.editor}
        contentEditable
        suppressContentEditableWarning
        onBlur={commit}
      />
    </div>
  );
}
