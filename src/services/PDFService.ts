import { PDFDocument } from 'pdf-lib';

export class PDFService {
    private static instance: PDFService;
    private pdfDocument: PDFDocument | null = null;

    private constructor() {}

    public static getInstance(): PDFService {
        if (!PDFService.instance) {
            PDFService.instance = new PDFService();
        }
        return PDFService.instance;
    }

    public async loadPDF(file: File): Promise<void> {
        try {
            const arrayBuffer = await file.arrayBuffer();
            this.pdfDocument = await PDFDocument.load(arrayBuffer);
        } catch (error) {
            console.error('Error loading PDF:', error);
            throw error;
        }
    }

    public async extractText(): Promise<string> {
        if (!this.pdfDocument) {
            throw new Error('No PDF document loaded');
        }

        // Note: pdf-lib doesn't support text extraction directly
        // We'll need to use a different library for text extraction
        // For now, we'll return a placeholder
        return 'PDF text extraction will be implemented';
    }

    public async getPageCount(): Promise<number> {
        if (!this.pdfDocument) {
            throw new Error('No PDF document loaded');
        }
        return this.pdfDocument.getPageCount();
    }

    public clearDocument(): void {
        this.pdfDocument = null;
    }
} 