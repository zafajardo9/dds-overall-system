
import { useState } from 'react';

// Helper to read file as base64 (used for fast path or fallback)
const readFileAsBase64 = (file: File): Promise<string> => {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = () => {
            const result = reader.result as string;
            // Handle cases where result might be ArrayBuffer (though readAsDataURL should return string)
            if (typeof result === 'string') {
                resolve(result.split(',')[1]);
            } else {
                reject(new Error("FileReader returned ArrayBuffer unexpectedly"));
            }
        };
        reader.onerror = () => reject(new Error("FileReader failed to read file"));
        reader.readAsDataURL(file);
    });
};

const sleep = (ms: number): Promise<void> =>
    new Promise((resolve) => setTimeout(resolve, ms));

// Robust Image Resizer
export const resizeImage = async (file: File, maxSize: number = 800): Promise<string> => {
    // 1. FAST PATH: If image is small (< 200KB), don't resize. Just return it.
    if (file.size < 200 * 1024) {
        return readFileAsBase64(file);
    }

    return new Promise((resolve, reject) => {
        // Validation
        if (!file.type.startsWith('image/')) {
            reject(new Error("File is not an image"));
            return;
        }

        const img = new Image();
        // Use ObjectURL instead of FileReader for loading the image source (Faster/Less Memory)
        const url = URL.createObjectURL(file);
        
        // Safety Timeout: If image takes > 2 seconds to decode, abort and try raw base64
        const timeoutId = setTimeout(() => {
            URL.revokeObjectURL(url);
            console.warn("Image resize timed out, falling back to raw file.");
            // Fallback to raw file if resize hangs
            readFileAsBase64(file).then(resolve).catch(reject);
        }, 2000);

        img.onload = () => {
            clearTimeout(timeoutId);
            
            try {
                const canvas = document.createElement('canvas');
                let width = img.width;
                let height = img.height;

                // Calculate Aspect Ratio
                if (width > height) {
                    if (width > maxSize) {
                        height *= maxSize / width;
                        width = maxSize;
                    }
                } else {
                    if (height > maxSize) {
                        width *= maxSize / height;
                        height = maxSize;
                    }
                }

                canvas.width = Math.floor(width);
                canvas.height = Math.floor(height);
                
                const ctx = canvas.getContext('2d');
                if (!ctx) {
                    throw new Error("Canvas context unavailable");
                }

                // White background for transparency safety
                ctx.fillStyle = "#FFFFFF";
                ctx.fillRect(0, 0, canvas.width, canvas.height);
                ctx.drawImage(img, 0, 0, canvas.width, canvas.height);
                
                // Export with lower quality for speed (0.6 is sufficient for OCR)
                const dataUrl = canvas.toDataURL('image/jpeg', 0.60);
                resolve(dataUrl.split(',')[1]); 
            } catch (err) {
                console.error("Canvas error, falling back to raw", err);
                readFileAsBase64(file).then(resolve).catch(reject);
            } finally {
                URL.revokeObjectURL(url);
            }
        };

        img.onerror = () => {
            clearTimeout(timeoutId);
            URL.revokeObjectURL(url);
            // Fallback to raw if image fails to load into Image object
            console.warn("Image load error, falling back to raw file.");
            readFileAsBase64(file).then(resolve).catch(reject);
        };

        img.src = url;
    });
};

// Main Hook
export const useFileProcessing = <T>(
    onProcess: (base64: string, mimeType: string, fileName: string) => Promise<T>, 
    acceptedTypes: string[] = ['application/pdf']
) => {
  const [isProcessing, setIsProcessing] = useState(false);
  const [progress, setProgress] = useState({ current: 0, total: 0 });
  const [error, setError] = useState<string | null>(null);

  const processFiles = async (
      files: File[], 
      onItemComplete?: (result: T) => void
  ): Promise<T[]> => {
    if (!files || files.length === 0) return [];

    setIsProcessing(true);
    setError(null);

    // Filter valid types
    const validFiles = files.filter(file => {
        return acceptedTypes.some(type => {
            if (type.endsWith('/*')) {
                const base = type.split('/')[0];
                return file.type.startsWith(base + '/');
            }
            return file.type === type;
        });
    });

    if (validFiles.length === 0) {
      setError(`Invalid file type(s).`);
      setIsProcessing(false);
      return [];
    }

    // Initialize Progress
    const totalFiles = validFiles.length;
    setProgress({ current: 0, total: totalFiles });
    let completedCount = 0;

    try {
        const successfulResults: T[] = [];
        // Keep parsing requests sequential to avoid bursting AI API quotas.
        const CONCURRENT_LIMIT = 1;
        const REQUEST_DELAY_MS = 1200;
        const queue = [...validFiles];
        
        const processSingleFile = async (file: File) => {
            try {
                let base64String = "";
                let mimeType = file.type;

                // Special handling for images in batch too
                if (file.type.startsWith('image/')) {
                    // Use robust resizer with optimized settings (800px)
                    base64String = await resizeImage(file, 800);
                    mimeType = 'image/jpeg'; 
                } else {
                    base64String = await readFileAsBase64(file);
                }

                if (!base64String) throw new Error("Empty content");

                const data = await onProcess(base64String, mimeType, file.name);
                
                if (onItemComplete) {
                    onItemComplete(data);
                }
                successfulResults.push(data);
            } catch (err) {
                console.error(`Error processing file ${file.name}:`, err);
            } finally {
                if (REQUEST_DELAY_MS > 0 && queue.length > 0) {
                    await sleep(REQUEST_DELAY_MS);
                }
                // Update progress regardless of success/failure
                completedCount++;
                setProgress({ current: completedCount, total: totalFiles });
            }
        };

        // Queue Execution
        const activeWorkers: Promise<void>[] = [];

        while (queue.length > 0 || activeWorkers.length > 0) {
            while (queue.length > 0 && activeWorkers.length < CONCURRENT_LIMIT) {
                const file = queue.shift();
                if (file) {
                    const promise = processSingleFile(file).then(() => {
                        const index = activeWorkers.indexOf(promise);
                        if (index > -1) activeWorkers.splice(index, 1);
                    });
                    activeWorkers.push(promise);
                }
            }
            if (activeWorkers.length > 0) {
                await Promise.race(activeWorkers);
            }
        }

        return successfulResults;

    } catch (err) {
        console.error("Batch processing error:", err);
        setError('Error processing files.');
        return [];
    } finally {
        setIsProcessing(false);
    }
  };
  
  const externalSetError = (message: string | null) => {
      setError(message);
      if (message === null) setIsProcessing(false);
  }

  return { processFiles, isProcessing, progress, error, setError: externalSetError };
};
