import { useRef, useState } from 'react';
import { Card, Text, makeStyles, tokens, Button, mergeClasses } from '@fluentui/react-components';
import { DocumentArrowUpFilled } from '@fluentui/react-icons';
import type { SourceFile } from '../types';
import { classify } from '../processing/pipeline';

const useStyles = makeStyles({
  zone: {
    borderTopWidth: '2px',
    borderRightWidth: '2px',
    borderBottomWidth: '2px',
    borderLeftWidth: '2px',
    borderTopStyle: 'dashed',
    borderRightStyle: 'dashed',
    borderBottomStyle: 'dashed',
    borderLeftStyle: 'dashed',
    borderTopColor: tokens.colorBrandStroke2,
    borderRightColor: tokens.colorBrandStroke2,
    borderBottomColor: tokens.colorBrandStroke2,
    borderLeftColor: tokens.colorBrandStroke2,
    borderRadius: tokens.borderRadiusXLarge,
    padding: tokens.spacingHorizontalXXL,
    textAlign: 'center',
    backgroundColor: tokens.colorBrandBackground2,
    transitionProperty: 'background-color',
    transitionDuration: tokens.durationFast,
    cursor: 'pointer',
  },
  zoneActive: {
    backgroundColor: tokens.colorBrandBackgroundSelected,
  },
  big: {
    fontSize: '36px',
    color: tokens.colorBrandForeground1,
  },
  hint: { color: tokens.colorNeutralForeground3, fontSize: tokens.fontSizeBase200 },
});

export interface LocalFilesDropZoneProps {
  onFiles: (files: Array<{ file: File; source: SourceFile }>) => void;
}

export function LocalFilesDropZone({ onFiles }: LocalFilesDropZoneProps) {
  const s = useStyles();
  const [active, setActive] = useState(false);
  const inputRef = useRef<HTMLInputElement>(null);

  function handle(list: FileList | File[] | null) {
    if (!list) return;
    const files: Array<{ file: File; source: SourceFile }> = [];
    for (const file of Array.from(list)) {
      const kind = classify(file.name);
      files.push({
        file,
        source: {
          id: `local-${file.name}-${file.size}-${file.lastModified}`,
          name: file.name,
          path: (file as File & { webkitRelativePath?: string }).webkitRelativePath || `/(local)/${file.name}`,
          size: file.size,
          modified: new Date(file.lastModified).toISOString(),
          kind,
        },
      });
    }
    onFiles(files);
  }

  return (
    <>
      <Card
        className={mergeClasses(s.zone, active && s.zoneActive)}
        onDragOver={event => { event.preventDefault(); setActive(true); }}
        onDragLeave={() => setActive(false)}
        onDrop={event => {
          event.preventDefault();
          setActive(false);
          handle(event.dataTransfer.files);
        }}
        onClick={() => inputRef.current?.click()}
        role="button"
        tabIndex={0}
      >
        <DocumentArrowUpFilled className={s.big} />
        <Text block weight="semibold" style={{ fontSize: tokens.fontSizeBase400, marginTop: tokens.spacingVerticalS }}>
          Drag &amp; drop files here
        </Text>
        <Text block className={s.hint}>
          .docx, .htm/.html, .pdf, .md · or click to pick files
        </Text>
        <div style={{ marginTop: tokens.spacingVerticalM, display: 'flex', gap: tokens.spacingHorizontalS, justifyContent: 'center', flexWrap: 'wrap' }}>
          <Button
            appearance="primary"
            icon={<DocumentArrowUpFilled />}
            onClick={event => { event.stopPropagation(); inputRef.current?.click(); }}
          >
            Choose files…
          </Button>
        </div>
      </Card>
      <input
        ref={inputRef}
        type="file"
        multiple
        accept=".docx,.html,.htm,.pdf,.md,.markdown"
        style={{ display: 'none' }}
        onChange={event => handle(event.target.files)}
      />
    </>
  );
}
