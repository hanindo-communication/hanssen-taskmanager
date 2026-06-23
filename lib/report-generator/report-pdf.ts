export async function elementToPdfBlob(el: HTMLElement): Promise<Blob> {
  const html2canvas = (await import('html2canvas')).default;
  const { jsPDF } = await import('jspdf');

  const canvas = await html2canvas(el, {
    scale: 2,
    useCORS: true,
    logging: false,
    onclone(doc) {
      doc.querySelectorAll('[data-no-pdf]').forEach((node) => {
        (node as HTMLElement).style.display = 'none';
      });
    },
  });

  const imgData = canvas.toDataURL('image/png');
  const pdf = new jsPDF({ orientation: 'p', unit: 'mm', format: 'a4' });
  const imgWidth = pdf.internal.pageSize.getWidth();
  const pageHeight = pdf.internal.pageSize.getHeight();
  const imgHeight = (canvas.height * imgWidth) / canvas.width;
  let heightLeft = imgHeight;
  let position = 0;

  pdf.addImage(imgData, 'PNG', 0, position, imgWidth, imgHeight);
  heightLeft -= pageHeight;

  while (heightLeft > 0) {
    position = heightLeft - imgHeight;
    pdf.addPage();
    pdf.addImage(imgData, 'PNG', 0, position, imgWidth, imgHeight);
    heightLeft -= pageHeight;
  }

  return pdf.output('blob');
}

export async function downloadElementAsPdf(el: HTMLElement, baseName: string): Promise<void> {
  const blob = await elementToPdfBlob(el);
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = `${baseName}.pdf`;
  a.click();
  URL.revokeObjectURL(url);
}

export async function downloadPdfZip(
  entries: Array<{ element: HTMLElement; fileName: string }>,
  zipBaseName: string
): Promise<void> {
  const JSZip = (await import('jszip')).default;
  const zip = new JSZip();

  for (const { element, fileName } of entries) {
    const blob = await elementToPdfBlob(element);
    zip.file(fileName, blob);
  }

  const zipBlob = await zip.generateAsync({ type: 'blob' });
  const url = URL.createObjectURL(zipBlob);
  const a = document.createElement('a');
  a.href = url;
  a.download = `${zipBaseName}.zip`;
  a.click();
  URL.revokeObjectURL(url);
}
