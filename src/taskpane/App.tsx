/// <reference types="office-js" />

import * as React from 'react';
import { useState, useEffect, useRef } from 'react';
import {
  Button,
  Textarea,
  Select,
  makeStyles,
  tokens,
  Label,
  Spinner
} from '@fluentui/react-components';
import mermaid from 'mermaid';
import './App.css';

const useStyles = makeStyles({
  root: {
    display: 'flex',
    flexDirection: 'column',
    height: '100vh',
    padding: '16px',
    paddingBottom: '60px',
    backgroundColor: tokens.colorNeutralBackground1,
    gap: '12px',
    overflow: 'auto'
  },
  header: {
    fontSize: '18px',
    fontWeight: 'bold',
    color: tokens.colorBrandForeground1,
    flexShrink: 0
  },
  editorSection: {
    display: 'flex',
    flexDirection: 'column',
    gap: '6px',
    minHeight: '150px',
    maxHeight: '200px',
    flexShrink: 0
  },
  textarea: {
    flex: 1,
    minHeight: '120px',
    fontFamily: 'monospace',
    fontSize: '12px'
  },
  previewSection: {
    display: 'flex',
    flexDirection: 'column',
    gap: '6px',
    flex: 1,
    minHeight: '150px',
    border: `1px solid ${tokens.colorNeutralStroke1}`,
    borderRadius: '4px',
    padding: '10px',
    backgroundColor: tokens.colorNeutralBackground2,
    overflow: 'auto'
  },
  buttonGroup: {
    display: 'flex',
    flexDirection: 'row',
    gap: '8px',
    flexWrap: 'wrap',
    justifyContent: 'space-between',
    flexShrink: 0,
    marginTop: '12px',
    marginBottom: '20px'
  },
  selectGroup: {
    display: 'flex',
    gap: '8px',
    alignItems: 'center',
    flexShrink: 0
  }
});

const defaultDiagrams = {
  flowchart: `flowchart TD
    A[Start] --> B{Is it?}
    B -->|Yes| C[OK]
    C --> D[Rethink]
    D --> B
    B ---->|No| E[End]`,
  sequence: `sequenceDiagram
    participant Alice
    participant Bob
    Alice->>John: Hello John, how are you?
    loop Healthcheck
        John->>John: Fight against hypochondria
    end
    Note right of John: Rational thoughts!
    John-->>Alice: Great!
    John->>Bob: How about you?
    Bob-->>John: Jolly good!`,
  class: `classDiagram
    Animal <|-- Duck
    Animal <|-- Fish
    Animal <|-- Zebra
    Animal : +int age
    Animal : +String gender
    Animal: +isMammal()
    Animal: +mate()
    class Duck{
      +String beakColor
      +swim()
      +quack()
    }
    class Fish{
      -int sizeInFeet
      -canEat()
    }
    class Zebra{
      +bool is_wild
      +run()
    }`
};

const App: React.FC = () => {
  const styles = useStyles();
  const [diagramType, setDiagramType] = useState<string>('flowchart');
  const [mermaidCode, setMermaidCode] = useState<string>(defaultDiagrams.flowchart);
  const [isRendering, setIsRendering] = useState<boolean>(false);
  const [error, setError] = useState<string>('');
  const previewRef = useRef<HTMLDivElement>(null);
  const svgRef = useRef<string>('');

  useEffect(() => {
    mermaid.initialize({
      startOnLoad: false,
      theme: 'default',
      securityLevel: 'loose',
      fontFamily: 'Arial, sans-serif'
    });
  }, []);

  useEffect(() => {
    renderDiagram();
  }, [mermaidCode]);

  const renderDiagram = async () => {
    if (!mermaidCode.trim()) {
      setError('Please enter Mermaid diagram code');
      return;
    }

    setIsRendering(true);
    setError('');

    try {
      const { svg } = await mermaid.render('preview-diagram', mermaidCode);
      svgRef.current = svg;
      
      if (previewRef.current) {
        previewRef.current.innerHTML = svg;
      }
    } catch (err: any) {
      setError(`Rendering error: ${err.message || 'Invalid Mermaid syntax'}`);
      if (previewRef.current) {
        previewRef.current.innerHTML = '';
      }
    } finally {
      setIsRendering(false);
    }
  };

  const handleDiagramTypeChange = (event: any, data: any) => {
    const newType = data.value;
    setDiagramType(newType);
    setMermaidCode(defaultDiagrams[newType as keyof typeof defaultDiagrams]);
  };

  const insertDiagramToSlide = async () => {
    if (!svgRef.current) {
      setError('Please render a valid diagram first');
      return;
    }

    try {
      const svgElement = previewRef.current?.querySelector('svg');
      if (!svgElement) {
        setError('No diagram to insert');
        return;
      }

      const clonedSvg = svgElement.cloneNode(true) as SVGElement;
      const bbox = svgElement.getBBox();
      const width = Math.ceil(bbox.width) || 800;
      const height = Math.ceil(bbox.height) || 600;
      
      clonedSvg.setAttribute('width', width.toString());
      clonedSvg.setAttribute('height', height.toString());
      clonedSvg.setAttribute('viewBox', `${bbox.x} ${bbox.y} ${bbox.width} ${bbox.height}`);

      const svgString = new XMLSerializer().serializeToString(clonedSvg);
      const canvas = document.createElement('canvas');
      const scale = 2;
      canvas.width = width * scale;
      canvas.height = height * scale;
      const ctx = canvas.getContext('2d');

      if (!ctx) {
        setError('Failed to create canvas context');
        return;
      }

      ctx.scale(scale, scale);
      const img = new Image();
      
      img.onload = () => {
        try {
          ctx.fillStyle = 'white';
          ctx.fillRect(0, 0, width, height);
          ctx.drawImage(img, 0, 0, width, height);

          // Try JPEG format instead of PNG for better compatibility
          canvas.toBlob((blob) => {
            if (!blob) {
              setError('Failed to create image blob');
              return;
            }

            const reader = new FileReader();
            reader.onloadend = () => {
              // Get base64 string without the data URI prefix
              const base64Full = reader.result as string;
              const base64data = base64Full.split(',')[1];
              
              // Try with just base64 string
              const Office = (window as any).Office;
              Office.context.document.setSelectedDataAsync(
                base64data,
                {
                  coercionType: Office.CoercionType.Image
                },
                (asyncResult: any) => {
                  if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                    // If that fails, try with full data URI
                    Office.context.document.setSelectedDataAsync(
                      base64Full,
                      {
                        coercionType: Office.CoercionType.Image
                      },
                      (asyncResult2: any) => {
                        if (asyncResult2.status === Office.AsyncResultStatus.Failed) {
                          setError(`Failed to insert: ${asyncResult2.error.message}. Try using Download PNG button instead.`);
                        } else {
                          setError('');
                          console.log('Diagram inserted successfully');
                        }
                      }
                    );
                  } else {
                    setError('');
                    console.log('Diagram inserted successfully');
                  }
                }
              );
            };
            reader.readAsDataURL(blob);
          }, 'image/jpeg', 0.95); // Try JPEG instead of PNG
          
        } catch (err: any) {
          setError(`Conversion error: ${err.message}`);
        }
      };

      img.onerror = () => {
        setError('Failed to load diagram image');
      };

      const svgDataUrl = 'data:image/svg+xml;charset=utf-8,' + encodeURIComponent(svgString);
      img.src = svgDataUrl;
      
    } catch (err: any) {
      setError(`Error: ${err.message}`);
      console.error('Insert error:', err);
    }
  };

  const downloadDiagram = () => {
    if (!svgRef.current) {
      setError('Please render a valid diagram first');
      return;
    }

    try {
      const svgElement = previewRef.current?.querySelector('svg');
      if (!svgElement) {
        setError('No diagram to download');
        return;
      }

      // Clone and prepare SVG
      const clonedSvg = svgElement.cloneNode(true) as SVGElement;
      const bbox = svgElement.getBBox();
      const width = Math.ceil(bbox.width) || 800;
      const height = Math.ceil(bbox.height) || 600;
      
      clonedSvg.setAttribute('width', width.toString());
      clonedSvg.setAttribute('height', height.toString());
      clonedSvg.setAttribute('viewBox', `${bbox.x} ${bbox.y} ${bbox.width} ${bbox.height}`);

      const svgString = new XMLSerializer().serializeToString(clonedSvg);
      const canvas = document.createElement('canvas');
      const scale = 2;
      canvas.width = width * scale;
      canvas.height = height * scale;
      const ctx = canvas.getContext('2d');

      if (!ctx) {
        setError('Failed to create canvas');
        return;
      }

      ctx.scale(scale, scale);
      const img = new Image();
      
      img.onload = () => {
        ctx.fillStyle = 'white';
        ctx.fillRect(0, 0, width, height);
        ctx.drawImage(img, 0, 0, width, height);

        canvas.toBlob((blob) => {
          if (!blob) {
            setError('Failed to create image');
            return;
          }

          const url = URL.createObjectURL(blob);
          const a = document.createElement('a');
          a.href = url;
          a.download = `mermaid-diagram-${Date.now()}.png`;
          document.body.appendChild(a);
          a.click();
          document.body.removeChild(a);
          URL.revokeObjectURL(url);
          
          console.log('Diagram downloaded successfully');
        }, 'image/png', 0.95);
      };

      img.onerror = () => {
        setError('Failed to load diagram');
      };

      const svgDataUrl = 'data:image/svg+xml;charset=utf-8,' + encodeURIComponent(svgString);
      img.src = svgDataUrl;
      
    } catch (err: any) {
      setError(`Download error: ${err.message}`);
    }
  };

  return (
    <div className={styles.root}>
      <div className={styles.header}>Mermaid Diagram Editor</div>
      
      <div className={styles.selectGroup}>
        <Label>Diagram Type:</Label>
        <Select value={diagramType} onChange={handleDiagramTypeChange}>
          <option value="flowchart">Flowchart</option>
          <option value="sequence">Sequence Diagram</option>
          <option value="class">Class Diagram</option>
        </Select>
      </div>

      <div className={styles.editorSection}>
        <Label>Mermaid Code:</Label>
        <Textarea
          className={styles.textarea}
          value={mermaidCode}
          onChange={(e, data) => setMermaidCode(data.value)}
          placeholder="Enter Mermaid diagram code here..."
          resize="vertical"
        />
      </div>

      <div className={styles.previewSection}>
        <Label>Preview:</Label>
        {isRendering && <Spinner label="Rendering diagram..." />}
        {error && <div style={{ color: 'red' }}>{error}</div>}
        <div ref={previewRef} style={{ textAlign: 'center' }} />
      </div>

      <div className={styles.buttonGroup}>
        <Button appearance="secondary" onClick={renderDiagram}>
          Refresh Preview
        </Button>
        <Button appearance="secondary" onClick={downloadDiagram}>
          Download PNG
        </Button>
        <Button appearance="primary" onClick={insertDiagramToSlide}>
          Insert into Slide
        </Button>
      </div>
    </div>
  );
};

export default App;

// Made with Bob
