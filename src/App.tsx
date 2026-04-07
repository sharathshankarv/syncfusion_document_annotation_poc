import { useEffect, useMemo, useRef, useState } from 'react';
import {
  Annotation,
  BookmarkView,
  Inject,
  LinkAnnotation,
  Magnification,
  Navigation,
  PdfViewerComponent,
  Print,
  TextSearch,
  TextSelection,
  ThumbnailView,
  Toolbar,
} from '@syncfusion/ej2-react-pdfviewer';
import {
  DocumentEditorContainerComponent,
  Toolbar as DocumentToolbar,
} from '@syncfusion/ej2-react-documenteditor';
import JSZip from 'jszip';
import './index.css';

DocumentEditorContainerComponent.Inject(DocumentToolbar);

const DOC_EDITOR_SERVICE_URL =
  'https://document.syncfusion.com/web-services/docx-editor/api/documenteditor/';
const PDF_RESOURCE_URL =
  'https://cdn.syncfusion.com/ej2/33.1.46/dist/ej2-pdfviewer-lib';

const STORAGE_KEYS = {
  activeTab: 'syncfusion-poc.activeTab',
  pdfName: 'syncfusion-poc.pdfName',
  pdfData: 'syncfusion-poc.pdfBase64',
  docxName: 'syncfusion-poc.docxName',
  docxSfdt: 'syncfusion-poc.docxSfdt',
  pptxName: 'syncfusion-poc.pptxName',
  pptxData: 'syncfusion-poc.pptxData',
};

type ActiveTab = 'pdf' | 'docx' | 'pptx';
type TextSelectionEndArgs = {
  selectedText?: string;
};

function App() {
  const [activeTab, setActiveTab] = useState<ActiveTab>(() => {
    const stored = localStorage.getItem(STORAGE_KEYS.activeTab) as ActiveTab | null;
    return stored ?? 'pdf';
  });

  const [pdfName, setPdfName] = useState<string>(() => localStorage.getItem(STORAGE_KEYS.pdfName) ?? '');
  const [pdfData, setPdfData] = useState<string>(() => localStorage.getItem(STORAGE_KEYS.pdfData) ?? '');

  const [docxName, setDocxName] = useState<string>(() => localStorage.getItem(STORAGE_KEYS.docxName) ?? '');

  const [pptxName, setPptxName] = useState<string>(() => localStorage.getItem(STORAGE_KEYS.pptxName) ?? '');
  const [pptxData, setPptxData] = useState<string>(() => localStorage.getItem(STORAGE_KEYS.pptxData) ?? '');
  const [pptxSlides, setPptxSlides] = useState<string[]>([]);

  const [statusMessage, setStatusMessage] = useState<string>('');
  const [selectedText, setSelectedText] = useState<string>('');
  const [pdfViewerKey, setPdfViewerKey] = useState<number>(0);
  const [highlightColor, setHighlightColor] = useState<string>('#FFE45C');

  const pdfInputRef = useRef<HTMLInputElement | null>(null);
  const docInputRef = useRef<HTMLInputElement | null>(null);
  const pptInputRef = useRef<HTMLInputElement | null>(null);
  const pdfViewerRef = useRef<PdfViewerComponent | null>(null);
  const docEditorRef = useRef<DocumentEditorContainerComponent | null>(null);

  useEffect(() => {
    localStorage.setItem(STORAGE_KEYS.activeTab, activeTab);
  }, [activeTab]);

  useEffect(() => {
    loadPdfInViewer();
  }, [pdfData, pdfName, pdfViewerKey]);

  useEffect(() => {
    const onResize = () => ensureDocxViewerHeight();
    onResize();
    window.addEventListener('resize', onResize);
    return () => window.removeEventListener('resize', onResize);
  }, []);

  const docEditorSummary = useMemo(() => {
    if (!docxName) {
      return 'No DOCX loaded yet';
    }
    return `Loaded: ${docxName}`;
  }, [docxName]);

  const openPdfPicker = () => pdfInputRef.current?.click();
  const openDocxPicker = () => docInputRef.current?.click();
  const openPptxPicker = () => pptInputRef.current?.click();

  const handlePdfUpload = async (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (!file) return;

    const base64 = await fileToBase64(file);
    setPdfName(file.name);
    setPdfData(base64);
    setPdfViewerKey((value) => value + 1);
    setStatusMessage(`Loaded PDF: ${file.name}`);

    localStorage.setItem(STORAGE_KEYS.pdfName, file.name);
    localStorage.setItem(STORAGE_KEYS.pdfData, base64);
    event.target.value = '';
  };

  const handleDocxUpload = async (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (!file) return;

    try {
      const endpoint = `${DOC_EDITOR_SERVICE_URL}Import`;
      const formData = new FormData();
      formData.append('files', file);

      const response = await fetch(endpoint, { method: 'POST', body: formData });
      if (!response.ok) {
        throw new Error(`DOCX import failed with status ${response.status}`);
      }

      const sfdt = await response.text();
      if (docEditorRef.current) {
        docEditorRef.current.documentEditor.open(sfdt);
        docEditorRef.current.documentEditor.documentName = trimExtension(file.name);
      }

      setDocxName(file.name);
      localStorage.setItem(STORAGE_KEYS.docxName, file.name);
      localStorage.setItem(STORAGE_KEYS.docxSfdt, sfdt);
      setStatusMessage(`Loaded DOCX: ${file.name}`);

      event.target.value = '';
      setActiveTab('docx');
    } catch (error) {
      setStatusMessage(
        `DOCX upload failed. Ensure internet access and Syncfusion service availability. ${
          error instanceof Error ? error.message : ''
        }`,
      );
    }
  };

  const handlePptxUpload = async (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (!file) return;

    try {
      const dataUrl = await fileToDataUrl(file);
      const slides = await extractPptxSlides(file);
      setPptxName(file.name);
      setPptxData(dataUrl);
      setPptxSlides(slides);
      localStorage.setItem(STORAGE_KEYS.pptxName, file.name);
      localStorage.setItem(STORAGE_KEYS.pptxData, dataUrl);
      setStatusMessage(`Loaded PPTX: ${file.name}`);
      event.target.value = '';
    } catch {
      setStatusMessage('PPTX upload worked, but slide preview extraction failed for this file.');
    }
  };

  const getPdfAnnotationStorageKey = (): string => {
    const name = pdfName || 'default-pdf';
    return `syncfusion-poc.pdfAnnotations.${name}`;
  };

  const persistPdfAnnotations = async () => {
    const viewer = pdfViewerRef.current as unknown as {
      exportAnnotationsAsObject?: () => Promise<string>;
    };
    if (!viewer?.exportAnnotationsAsObject) return;

    try {
      const annotationObject = await viewer.exportAnnotationsAsObject();
      localStorage.setItem(getPdfAnnotationStorageKey(), annotationObject);
    } catch {
      // no-op
    }
  };

  const restorePdfAnnotations = () => {
    const stored = localStorage.getItem(getPdfAnnotationStorageKey());
    if (!stored) return;

    const viewer = pdfViewerRef.current as unknown as {
      importAnnotation?: (value: unknown) => void;
    };
    if (!viewer?.importAnnotation) return;

    try {
      viewer.importAnnotation(JSON.parse(stored));
      setStatusMessage(`Restored saved annotations for ${pdfName || 'PDF'}.`);
    } catch {
      setStatusMessage('Could not restore saved annotations from local storage.');
    }
  };

  const setPdfAnnotationMode = (mode: 'Highlight' | 'Underline' | 'Strikethrough' | 'None') => {
    const viewer = pdfViewerRef.current as unknown as {
      annotationModule?: { setAnnotationMode?: (annotationType: string) => void };
      annotation?: { setAnnotationMode?: (annotationType: string) => void };
    };

    viewer?.annotationModule?.setAnnotationMode?.(mode);
    viewer?.annotation?.setAnnotationMode?.(mode);

    setStatusMessage(
      mode === 'None' ? 'Annotation mode cleared.' : `${mode} mode enabled. Select text in the PDF to annotate.`,
    );
  };

  const applyHighlightColor = (color: string) => {
    setHighlightColor(color);
    setPdfAnnotationMode('Highlight');
    setStatusMessage(`Highlight color set to ${color}.`);
  };

  const onPdfTextSelectionEnd = (args: TextSelectionEndArgs) => {
    const text = (args.selectedText ?? '').trim();
    if (!text) return;

    setSelectedText(text);
    console.log('Selected text:', text);
  };

  const onPdfViewerCreated = () => {
    loadPdfInViewer();
  };

  const loadPdfInViewer = () => {
    if (!pdfData || !pdfViewerRef.current) return;
    setTimeout(() => {
      if (pdfViewerRef.current && pdfData) {
        pdfViewerRef.current.load(pdfData, pdfName || 'uploaded.pdf');
      }
    }, 0);
  };

  const clearPdfState = () => {
    const annotationKey = getPdfAnnotationStorageKey();
    setPdfName('');
    setPdfData('');
    setSelectedText('');
    setPdfViewerKey((value) => value + 1);
    localStorage.removeItem(STORAGE_KEYS.pdfName);
    localStorage.removeItem(STORAGE_KEYS.pdfData);
    localStorage.removeItem(annotationKey);
    setStatusMessage('PDF file discarded.');
  };

  const clearDocxState = () => {
    setDocxName('');
    localStorage.removeItem(STORAGE_KEYS.docxName);
    localStorage.removeItem(STORAGE_KEYS.docxSfdt);
    const editor = docEditorRef.current?.documentEditor as unknown as { openBlank?: () => void };
    editor?.openBlank?.();
    setStatusMessage('DOCX file discarded.');
  };

  const clearPptxState = () => {
    setPptxName('');
    setPptxData('');
    setPptxSlides([]);
    localStorage.removeItem(STORAGE_KEYS.pptxName);
    localStorage.removeItem(STORAGE_KEYS.pptxData);
    setStatusMessage('PPTX file discarded.');
  };

  const handleTabChange = (nextTab: ActiveTab) => {
    if (nextTab === activeTab) return;

    if (activeTab === 'pdf') clearPdfState();
    if (activeTab === 'docx') clearDocxState();
    if (activeTab === 'pptx') clearPptxState();

    setActiveTab(nextTab);
  };

  const ensureDocxViewerHeight = () => {
    const viewerContainer = document.getElementById('docx-editor_editor_viewerContainer');
    if (viewerContainer) {
      viewerContainer.style.height = 'calc(100vh - 250px)';
      viewerContainer.style.minHeight = '640px';
    }
  };

  const handleDocEditorCreated = () => {
    ensureDocxViewerHeight();
    const cached = localStorage.getItem(STORAGE_KEYS.docxSfdt);
    if (cached && docEditorRef.current) {
      docEditorRef.current.documentEditor.open(cached);
    }
  };

  const saveDocxDraftToStorage = () => {
    const editor = docEditorRef.current?.documentEditor as unknown as { serialize?: () => string };
    if (editor?.serialize) {
      localStorage.setItem(STORAGE_KEYS.docxSfdt, editor.serialize());
    }
  };

  const preventDocxTyping = (event: React.KeyboardEvent<HTMLDivElement>) => {
    const target = event.target as HTMLElement | null;
    const isCommentOrDialogInput =
      !!target &&
      (target.tagName === 'INPUT' ||
        target.tagName === 'TEXTAREA' ||
        target.isContentEditable ||
        !!target.closest('.e-dialog') ||
        !!target.closest('.e-de-cmt-sub-container') ||
        !!target.closest('.e-input'));
    if (isCommentOrDialogInput) return;

    const blocked = ['Backspace', 'Delete', 'Enter', 'Tab'];
    if (event.key.length === 1 || blocked.includes(event.key)) {
      event.preventDefault();
      setStatusMessage('Direct typing is disabled. Use annotation actions only.');
    }
  };

  const blockDocxPaste = (event: React.ClipboardEvent<HTMLDivElement>) => {
    const target = event.target as HTMLElement | null;
    const isCommentOrDialogInput =
      !!target &&
      (target.tagName === 'INPUT' ||
        target.tagName === 'TEXTAREA' ||
        target.isContentEditable ||
        !!target.closest('.e-dialog') ||
        !!target.closest('.e-de-cmt-sub-container') ||
        !!target.closest('.e-input'));
    if (isCommentOrDialogInput) return;

    event.preventDefault();
    setStatusMessage('Pasting content is disabled in annotation-only mode.');
  };

  const onDocxSelectionChange = () => {
    const editor = docEditorRef.current?.documentEditor as unknown as { selection?: { text?: string } };
    const text = editor?.selection?.text?.trim();
    if (text) {
      console.log('Selected text:', text);
    }
  };

  const highlightDocxSelection = () => {
    const editor = docEditorRef.current?.documentEditor as unknown as {
      selection?: { text?: string; characterFormat?: { highlightColor?: string } };
      editor?: { applyHighlightColor?: (color: string) => void };
    };

    if (!editor?.selection?.text?.trim()) {
      setStatusMessage('Select text first, then click Highlight Selection.');
      return;
    }

    if (editor?.editor?.applyHighlightColor) {
      editor.editor.applyHighlightColor('Yellow');
    } else if (editor?.selection?.characterFormat) {
      editor.selection.characterFormat.highlightColor = 'Yellow';
    }

    saveDocxDraftToStorage();
    setStatusMessage('Applied yellow highlight to selected DOCX text.');
  };

  const addDocxComment = () => {
    const editor = docEditorRef.current?.documentEditor as unknown as {
      selection?: { text?: string };
      editor?: { insertComment?: (text: string) => void };
    };

    if (!editor?.selection?.text?.trim()) {
      setStatusMessage('Select text first, then click Add Comment.');
      return;
    }

    if (editor?.editor?.insertComment) {
      editor.editor.insertComment('Review note');
      saveDocxDraftToStorage();
      setStatusMessage('Comment annotation added to selected DOCX text.');
      return;
    }

    setStatusMessage('Comment API is unavailable in this current editor mode.');
  };

  const exportAnnotatedPdf = () => {
    const viewer = pdfViewerRef.current as unknown as {
      saveAsBlob?: () => Promise<Blob>;
      download?: () => void;
    };
    if (!viewer) return;

    if (viewer.saveAsBlob) {
      viewer
        .saveAsBlob()
        .then((blob) => {
          const url = URL.createObjectURL(blob);
          const link = document.createElement('a');
          link.href = url;
          link.download = pdfName || 'annotated.pdf';
          link.click();
          URL.revokeObjectURL(url);
          setStatusMessage(`Downloaded ${link.download}`);
        })
        .catch(() => {
          viewer.download?.();
        });
      return;
    }

    viewer.download?.();
  };

  const exportAnnotatedDocx = () => {
    const docEditor = docEditorRef.current?.documentEditor;
    if (!docEditor) return;

    const fileName = trimExtension(docxName || 'annotated-document');
    docEditor.save(fileName, 'Docx');
    setStatusMessage(`Export started for ${fileName}.docx`);
  };

  const exportPptxFile = () => {
    if (!pptxData) return;

    const link = document.createElement('a');
    link.href = pptxData;
    link.download = pptxName || 'slides.pptx';
    link.click();
    setStatusMessage(`Downloaded ${link.download}`);
  };

  return (
    <main className="app-shell">
      <header className="topbar">
        <div>
          <h1>Syncfusion Document Review POC</h1>
          <p>PDF and DOCX support annotations. PPTX is read-only text preview.</p>
        </div>
        <div className="badge-row">
          <span className="badge">Frontend only</span>
          <span className="badge">Local storage enabled</span>
        </div>
      </header>

      <section className="tab-row" aria-label="Document type tabs">
        <button className={activeTab === 'pdf' ? 'tab active' : 'tab'} onClick={() => handleTabChange('pdf')} type="button">
          PDF
        </button>
        <button className={activeTab === 'docx' ? 'tab active' : 'tab'} onClick={() => handleTabChange('docx')} type="button">
          DOCX
        </button>
        <button className={activeTab === 'pptx' ? 'tab active' : 'tab'} onClick={() => handleTabChange('pptx')} type="button">
          PPTX
        </button>
      </section>

      {statusMessage && <section className="status">{statusMessage}</section>}

      {activeTab === 'pdf' && (
        <section className="panel">
          <div className="toolbar-row">
            <input ref={pdfInputRef} type="file" accept=".pdf,application/pdf" onChange={handlePdfUpload} className="hidden-input" />
            <button className="action" onClick={openPdfPicker} type="button">Upload PDF</button>
            <button className="action ghost" onClick={exportAnnotatedPdf} type="button">Download Annotated PDF</button>
            <button className="action ghost" onClick={clearPdfState} type="button">Discard File</button>
            <button className="action ghost" onClick={() => setPdfAnnotationMode('Highlight')} type="button">Highlight Mode</button>
            <button className="action ghost" onClick={() => setPdfAnnotationMode('Underline')} type="button">Underline Mode</button>
            <button className="action ghost" onClick={() => setPdfAnnotationMode('Strikethrough')} type="button">Strike Mode</button>
            <button className="action ghost" onClick={() => setPdfAnnotationMode('None')} type="button">Normal Mode</button>
            <span className="meta">{pdfName || 'No PDF selected'}</span>
          </div>
          {selectedText && <p className="hint">Selected text: {selectedText}</p>}

          <div className="toolbar-row">
            <span className="meta">Highlight Color:</span>
            <button className="action ghost" onClick={() => applyHighlightColor('#FFE45C')} type="button">Yellow</button>
            <button className="action ghost" onClick={() => applyHighlightColor('#B9F18A')} type="button">Green</button>
            <button className="action ghost" onClick={() => applyHighlightColor('#FFC3D9')} type="button">Pink</button>
            <button className="action ghost" onClick={() => applyHighlightColor('#A8D8FF')} type="button">Blue</button>
            <input className="color-input" type="color" value={highlightColor} onChange={(event) => applyHighlightColor(event.target.value)} />
          </div>

          <div className="viewer-wrap">
            {pdfData ? (
              <PdfViewerComponent
                id="pdf-viewer"
                key={pdfViewerKey}
                ref={pdfViewerRef}
                resourceUrl={PDF_RESOURCE_URL}
                enableTextSelection
                created={onPdfViewerCreated}
                textSelectionEnd={onPdfTextSelectionEnd}
                documentLoad={restorePdfAnnotations}
                annotationAdd={persistPdfAnnotations}
                annotationRemove={persistPdfAnnotations}
                annotationMove={persistPdfAnnotations}
                annotationResize={persistPdfAnnotations}
                annotationPropertiesChange={persistPdfAnnotations}
                highlightSettings={{ color: highlightColor, opacity: 0.6 }}
                style={{ height: '100%', width: '100%' }}
              >
                <Inject
                  services={[
                    Toolbar,
                    Magnification,
                    Navigation,
                    LinkAnnotation,
                    BookmarkView,
                    ThumbnailView,
                    Print,
                    TextSelection,
                    Annotation,
                    TextSearch,
                  ]}
                />
              </PdfViewerComponent>
            ) : (
              <div className="empty-state">Upload a PDF, use text markup tools, then download.</div>
            )}
          </div>
        </section>
      )}

      {activeTab === 'docx' && (
        <section className="panel">
          <div className="toolbar-row">
            <input
              ref={docInputRef}
              type="file"
              accept=".docx,application/vnd.openxmlformats-officedocument.wordprocessingml.document"
              onChange={handleDocxUpload}
              className="hidden-input"
            />
            <button className="action" onClick={openDocxPicker} type="button">Upload DOCX</button>
            <button className="action ghost" onClick={exportAnnotatedDocx} type="button">Download Annotated DOCX</button>
            <button className="action ghost" onClick={clearDocxState} type="button">Discard File</button>
            <button className="action ghost" onClick={highlightDocxSelection} type="button">Highlight Selection</button>
            <button className="action ghost" onClick={addDocxComment} type="button">Add Comment</button>
            <span className="meta">{docEditorSummary}</span>
          </div>

          <div className="viewer-wrap" onKeyDownCapture={preventDocxTyping} onPasteCapture={blockDocxPaste}>
            <DocumentEditorContainerComponent
              id="docx-editor"
              ref={docEditorRef}
              height="100%"
              serviceUrl={DOC_EDITOR_SERVICE_URL}
              enableToolbar={false}
              created={handleDocEditorCreated}
              contentChange={saveDocxDraftToStorage}
              selectionChange={onDocxSelectionChange}
            />
          </div>

          <p className="hint">DOCX is annotation-only: direct typing is blocked, but highlight/comment annotations are allowed.</p>
        </section>
      )}

      {activeTab === 'pptx' && (
        <section className="panel">
          <div className="toolbar-row">
            <input
              ref={pptInputRef}
              type="file"
              accept=".pptx,application/vnd.openxmlformats-officedocument.presentationml.presentation"
              onChange={handlePptxUpload}
              className="hidden-input"
            />
            <button className="action" onClick={openPptxPicker} type="button">Upload PPTX</button>
            <button className="action ghost" onClick={exportPptxFile} type="button">Download PPTX</button>
            <button className="action ghost" onClick={clearPptxState} type="button">Discard File</button>
            <span className="meta">{pptxName || 'No PPTX selected'}</span>
          </div>

          <div className="empty-state">
            {pptxSlides.length > 0 ? (
              <div className="pptx-preview">
                <h3>Slide Text Preview</h3>
                {pptxSlides.map((slide, index) => (
                  <div className="slide-card" key={`slide-${index + 1}`}>
                    <strong>Slide {index + 1}</strong>
                    <p>{slide || 'No text extracted from this slide.'}</p>
                  </div>
                ))}
              </div>
            ) : (
              'Upload a PPTX to preview extracted slide text in read-only mode.'
            )}
          </div>
        </section>
      )}
    </main>
  );
}

function fileToBase64(file: File): Promise<string> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => {
      const value = String(reader.result ?? '');
      const marker = 'base64,';
      const index = value.indexOf(marker);
      if (index === -1) {
        reject(new Error('Could not parse PDF as base64.'));
        return;
      }

      resolve(value.substring(index + marker.length));
    };
    reader.onerror = () => reject(reader.error);
    reader.readAsDataURL(file);
  });
}

function fileToDataUrl(file: File): Promise<string> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(String(reader.result ?? ''));
    reader.onerror = () => reject(reader.error);
    reader.readAsDataURL(file);
  });
}

function trimExtension(name: string): string {
  const extensionIndex = name.lastIndexOf('.');
  if (extensionIndex === -1) return name;
  return name.substring(0, extensionIndex);
}

async function extractPptxSlides(file: File): Promise<string[]> {
  const zip = await JSZip.loadAsync(await file.arrayBuffer());
  const slideFiles = Object.keys(zip.files)
    .filter((key) => /^ppt\/slides\/slide\d+\.xml$/.test(key))
    .sort((a, b) => {
      const aNumber = Number((a.match(/slide(\d+)\.xml$/) ?? [])[1] ?? 0);
      const bNumber = Number((b.match(/slide(\d+)\.xml$/) ?? [])[1] ?? 0);
      return aNumber - bNumber;
    });

  const slides: string[] = [];
  for (const key of slideFiles) {
    const xml = await zip.files[key].async('string');
    slides.push(extractTextFromSlideXml(xml));
  }

  return slides;
}

function extractTextFromSlideXml(xml: string): string {
  const parser = new DOMParser();
  const doc = parser.parseFromString(xml, 'application/xml');
  const nodes = Array.from(doc.getElementsByTagName('a:t'));
  return nodes
    .map((node) => node.textContent?.trim() ?? '')
    .filter((value) => value.length > 0)
    .join(' ');
}

export default App;
