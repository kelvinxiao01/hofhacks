import * as pdfjsLib from 'pdfjs-dist';

// Set the worker source path
pdfjsLib.GlobalWorkerOptions.workerSrc = `https://cdnjs.cloudflare.com/ajax/libs/pdf.js/${pdfjsLib.version}/pdf.worker.min.js`;

export interface PDFContent {
  text: string;
  numPages: number;
  metadata: {
    title?: string;
    author?: string;
    subject?: string;
    keywords?: string;
    creationDate?: string;
    modificationDate?: string;
  };
}

interface PDFMetadata {
  info: {
    Title?: string;
    Author?: string;
    Subject?: string;
    Keywords?: string;
    CreationDate?: string;
    ModDate?: string;
    [key: string]: any;
  };
}

export class PDFService {
  static async parsePDF(file: File): Promise<PDFContent> {
    try {
      // Convert File to ArrayBuffer
      const arrayBuffer = await file.arrayBuffer();
      
      // Load the PDF document
      const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
      
      // Extract text from all pages
      let fullText = '';
      for (let i = 1; i <= pdf.numPages; i++) {
        const page = await pdf.getPage(i);
        const textContent = await page.getTextContent();
        const pageText = textContent.items.map((item: any) => item.str).join(' ');
        fullText += pageText + '\n';
      }
      
      // Get document info
      const info = await pdf.getMetadata() as PDFMetadata;
      
      // Log the extracted text for development
      console.log('Extracted PDF Text:', fullText);
      console.log('Number of Pages:', pdf.numPages);
      console.log('PDF Metadata:', info);

      return {
        text: fullText,
        numPages: pdf.numPages,
        metadata: {
          title: info?.info?.Title,
          author: info?.info?.Author,
          subject: info?.info?.Subject,
          keywords: info?.info?.Keywords,
          creationDate: info?.info?.CreationDate,
          modificationDate: info?.info?.ModDate,
        }
      };
    } catch (error) {
      console.error('Error parsing PDF:', error);
      throw new Error('Failed to parse PDF file');
    }
  }
} 