
function buildExecutiveSummaryData({
    situationTexts,
    emotionTexts,
    needTexts,
    actionTexts,
    customSituationTexts,
    customActionTexts,
    notes,
    context
}) {
    const totalSituations = situationTexts.length;
    const totalEmotions = emotionTexts.length;
    const totalNeeds = needTexts.length;
    const totalActions = actionTexts.length;

    const totalCustomSituations = customSituationTexts.length;
    const totalCustomActions = customActionTexts.length;
    const totalCustom = totalCustomSituations + totalCustomActions;

    const totalMapItems =
        totalSituations + totalEmotions + totalNeeds + totalActions;

    const hasNotes = Boolean(notes && notes.trim());

    return {
        totalSituations,
        totalEmotions,
        totalNeeds,
        totalActions,
        totalCustomSituations,
        totalCustomActions,
        totalCustom,
        totalMapItems,
        hasNotes,
        context: context || 'Plantilla utilitzada: En blanc',

        distributionRows: [
    ['Situacions', String(totalSituations)],
    ['Emocions', String(totalEmotions)],
    ['Necessitats', String(totalNeeds)],
    ['Accions', String(totalActions)]
],

recordRows: [
    ['Plantilla', context || 'En blanc'],
    /*['Situacions registrades', totalSituations > 0 ? 'Sí' : 'No'],
    ['Emocions registrades', totalEmotions > 0 ? 'Sí' : 'No'],
    ['Necessitats registrades', totalNeeds > 0 ? 'Sí' : 'No'],
    ['Accions registrades', totalActions > 0 ? 'Sí' : 'No'],*/
    ['Observacions complementàries', hasNotes ? 'Sí' : 'No']
],

customRows: [
    ['Situacions noves incorporades', String(totalCustomSituations)],
    ['Accions noves incorporades', String(totalCustomActions)],
    ['Total d’elements personalitzats', String(totalCustom)]
],

useLines: [
    'Es pot ampliar el registre de situacions.',
    'Es pot incorporar més detall en necessitats o accions.',
    'Es pot desenvolupar el contingut del mapa en sessions posteriors.'
]
    };
}

function drawExecutiveSummaryPage_OLD(doc, {
    logoDataUrl,
    generatedAt,
    title,
    subtitle,
    people,
    notes,
    summary,
    pageNumber,
    THEME
}) {
    const w = doc.internal.pageSize.getWidth();
    const h = doc.internal.pageSize.getHeight();
    const m = 14;
    const usableW = w - (m * 2);

    const headerBottom = drawPageHeader(doc, {
        logoDataUrl,
        title: 'Resum executiu',
        subtitle: title,
        pageWidth: w,
        margin: m
    });

    let y = headerBottom + 10;

    setFont(doc, 'bold', 15, THEME.ink);
    doc.text('Lectura general de la sessió', m, y);
    y += 7;

    setFont(doc, 'normal', 10.5, THEME.text);
    const introLines = doc.splitTextToSize(
        summary.summaryLead + ' ' + summary.reading,
        usableW
    );
    doc.text(introLines, m, y);
    y += (introLines.length * 5.2) + 6;

    // Caja de síntesis
    drawSectionCard(doc, {
        x: m,
        y,
        w: usableW,
        h: 34,
        title: 'Síntesi executiva',
        fillColor: THEME.softBg,
        accentColor: THEME.brandBlue
    });

    setFont(doc, 'normal', 9.5, THEME.text);
    doc.text(`Zona relacional principal: ${summary.relationalZone}`, m + 7, y + 15);
    doc.text(`Moment de treball: ${summary.stageZone}`, m + 7, y + 21);
    doc.text(`Persones participants: ${people || 'No indicades'}`, m + 7, y + 27);
    y += 42;

    // Indicadores
    const statW = (usableW - 9) / 4;
    const statData = [
        ['Situacions', summary.totalSituations, THEME.brandRed],
        ['Emocions', summary.totalEmotions, THEME.brandViolet],
        ['Necessitats', summary.totalNeeds, THEME.brandBlue],
        ['Accions', summary.totalActions, THEME.brandOrange]
    ];

    statData.forEach((item, index) => {
        const x = m + (index * (statW + 3));
        drawRoundedCard(doc, x, y, statW, 22, THEME.white, THEME.line, 3);
        doc.setFillColor(...item[2]);
        doc.roundedRect(x, y, statW, 4, 3, 3, 'F');

        setFont(doc, 'bold', 12, THEME.ink);
        doc.text(String(item[1]), x + 6, y + 12);

        setFont(doc, 'normal', 8.5, THEME.muted);
        doc.text(item[0], x + 6, y + 18);
    });

    y += 30;

    // Fortalezas
    const strengthsH = Math.max(34, estimateListHeight(doc, summary.strengths, usableW - 12, 8) + 14);
    drawSectionCard(doc, {
        x: m,
        y,
        w: usableW,
        h: strengthsH,
        title: 'Elements que ja sostenen l’avanç',
        fillColor: THEME.needBg,
        accentColor: THEME.brandGreen
    });
    drawListInCard(doc, m + 7, y + 15, summary.strengths, usableW - 12, 8);
    y += strengthsH + 8;

    // Alertas / foco
    const cautionsH = Math.max(34, estimateListHeight(doc, summary.cautions, usableW - 12, 8) + 14);
    drawSectionCard(doc, {
        x: m,
        y,
        w: usableW,
        h: cautionsH,
        title: 'Aspectes a prioritzar',
        fillColor: THEME.situationBg,
        accentColor: THEME.brandRed
    });
    drawListInCard(doc, m + 7, y + 15, summary.cautions, usableW - 12, 8);
    y += cautionsH + 8;

    // Observaciones
    if (notes) {
        const notesLines = doc.splitTextToSize(notes, usableW - 12).slice(0, 5);
        const notesH = Math.max(28, (notesLines.length * 5) + 16);

        drawSectionCard(doc, {
            x: m,
            y,
            w: usableW,
            h: notesH,
            title: 'Observacions de context',
            fillColor: THEME.white,
            accentColor: THEME.brandViolet
        });

        setFont(doc, 'normal', 9.2, THEME.text);
        doc.text(notesLines, m + 7, y + 15);
    }

    drawPageFooter(doc, {
        generatedAt,
        pageNumber,
        pageWidth: w,
        pageHeight: h,
        margin: m
    });
}


function drawExecutiveSummaryPage(doc, {
    logoDataUrl,
    generatedAt,
    title,
    subtitle,
    people,
    notes,
    summary,
    pageNumber,
    THEME
}) {
    const w = doc.internal.pageSize.getWidth();
    const h = doc.internal.pageSize.getHeight();
    const m = 16;
    const usableW = w - (m * 2);

    let y = drawMonoHeader(doc, {
        title: 'Resum',
        subtitle: title || '',
        pageWidth: w,
        margin: m
    });

    // 1. Datos generales
    drawMonoSectionTitle(doc, '1. Dades generals', m, y, usableW);
    y += 8;

    doc.setFont('helvetica', 'normal');
    doc.setFontSize(9.2);
    doc.setTextColor(40, 40, 40);
    doc.text(`Data de generació: ${generatedAt}`, m, y);
    y += 5.5;
    doc.text(`Persones participants: ${people || 'No consignades'}`, m, y);
    y += 10;

    // 2. Resumen cuantitativo
    drawMonoSectionTitle(doc, '2. Resum quantitatiu', m, y, usableW);
    y += 8;

    y = drawMonoTable(doc, {
        x: m,
        y,
        w: usableW,
        header: ['Categoría', 'Total'],
        rows: summary.distributionRows,
        colRatios: [0.72, 0.28],
        rowH: 8
    });

    y += 8;

    // 3. Registro complementario
    drawMonoSectionTitle(doc, '3. Altres dades', m, y, usableW);
    y += 8;

    y = drawMonoTable(doc, {
        x: m,
        y,
        w: usableW,
        rows: summary.recordRows,
        colRatios: [0.72, 0.28],
        rowH: 8
    });

    y += 8;

    y = drawMonoTable(doc, {
        x: m,
        y,
        w: usableW,
        rows: summary.customRows,
        colRatios: [0.72, 0.28],
        rowH: 8
    });

    y += 8;

    // 4. Posibilidades de uso
    /*drawMonoSectionTitle(doc, '4. Posibilidades de uso', m, y, usableW);*/
    /*y += 8;

    y = drawMonoLines(doc, {
        x: m,
        y,
        w: usableW,
        lines: summary.useLines.map(line => `- ${line}`),
        lineHeight: 4.8
    });*/

    // 5. Observaciones consignadas
    if (notes && notes.trim()) {
        y += 6;
        drawMonoSectionTitle(doc, '4. Observacions consignades', m, y, usableW);
        y += 8;

        doc.setDrawColor(170, 170, 170);
        doc.setLineWidth(0.2);
        doc.rect(m, y - 3, usableW, 28);

        doc.setFont('helvetica', 'normal');
        doc.setFontSize(9);
        doc.setTextColor(40, 40, 40);

        const noteLines = doc.splitTextToSize(notes, usableW - 6).slice(0, 6);
        doc.text(noteLines, m + 3, y + 2);
    }

    drawPageFooter(doc, {
        generatedAt,
        pageNumber,
        pageWidth: w,
        pageHeight: h,
        margin: m
    });
}

function drawPdfLogo(doc, logoDataUrl, x, y, targetW = 30) {
    if (!logoDataUrl) {
        console.warn('Logo no disponible: logoDataUrl es null');
        return null;
    }

    try {
        const props = doc.getImageProperties(logoDataUrl);
        const aspectRatio = props.width / props.height;
        const targetH = targetW / aspectRatio;

        doc.addImage(logoDataUrl, 'PNG', x, y, targetW, targetH);
        return { width: targetW, height: targetH };
    } catch (e) {
        console.warn('No s’ha pogut dibuixar el logo al PDF', e);
        return null;
    }
}

function getBoardTagTexts(areaId) {
    const area = document.getElementById(areaId);
    if (!area) return [];

    return Array.from(area.querySelectorAll('.organizer-tag.on-board'))
        .map(tag => (tag.dataset.text || '').trim())
        .filter(Boolean)
        .map(text => text.replace(/\s+/g, ' ').trim());
}

function cleanExportTagText(text) {
    return String(text)
        .replace(/^📷\s*/, '')
        .replace(/^🌼\s*/, '')
        .replace(/[\u{1F300}-\u{1FAFF}]/gu, '')
        .replace(/[\u2600-\u27BF]/g, '')
        .replace(/\s+/g, ' ')
        .trim();
}

function drawPageHeader(doc, {
  logoDataUrl,
  title,
  subtitle = '',
  pageWidth,
  margin = 14,
  color = [248, 250, 252],
  showLogo = true
}) {
  const headerY = 10;
  const headerH = 22;
  const contentW = pageWidth - margin * 2;

  doc.setFillColor(...color);
  doc.roundedRect(margin, headerY, contentW, headerH, 4, 4, 'F');
  doc.setDrawColor(226, 232, 240);
  doc.roundedRect(margin, headerY, contentW, headerH, 4, 4, 'S');

  doc.setFont('helvetica', 'bold');
  doc.setFontSize(15);
  doc.setTextColor(24, 24, 27);
  doc.text(title, margin + 4, headerY + 8);

  if (subtitle) {
    doc.setFont('helvetica', 'normal');
    doc.setFontSize(9);
    doc.setTextColor(100, 116, 139);
    doc.text(subtitle, margin + 4, headerY + 15);
  }

  if (showLogo && logoDataUrl) {
    const logoW = 28;
    const logoH = 10;
    const logoX = pageWidth - margin - logoW - 2;
    const logoY = headerY + 3.5;
    doc.addImage(logoDataUrl, 'PNG', logoX, logoY, logoW, logoH);
  }

  return headerY + headerH;
}

function drawPageFooter(doc, {
  generatedAt,
  pageNumber,
  pageWidth,
  pageHeight,
  margin = 14
}) {
  const y = pageHeight - 8;

  doc.setDrawColor(226, 232, 240);
  doc.line(margin, y - 4, pageWidth - margin, y - 4);

  doc.setFont('helvetica', 'normal');
  doc.setFontSize(8);
  doc.setTextColor(100, 116, 139);
  doc.text(`Generat: ${generatedAt}`, margin, y);

  doc.text(`Pàgina ${pageNumber}`, pageWidth - margin, y, { align: 'right' });
}

function drawSectionCard(doc, {
  x, y, w, h,
  title,
  fillColor = [255, 255, 255],
  accentColor = [220, 38, 38]
}) {
  doc.setFillColor(...fillColor);
  doc.roundedRect(x, y, w, h, 4, 4, 'F');
  doc.setDrawColor(226, 232, 240);
  doc.roundedRect(x, y, w, h, 4, 4, 'S');

  doc.setFillColor(...accentColor);
  doc.roundedRect(x, y, 3, h, 2, 2, 'F');

  doc.setFont('helvetica', 'bold');
  doc.setFontSize(11);
  doc.setTextColor(24, 24, 27);
  doc.text(title, x + 7, y + 8);
}

function drawCoverPage_OLD(doc, {
    logoDataUrl,
    title,
    subtitle,
    people,
    notes,
    generatedAt
}) {
    const w = doc.internal.pageSize.getWidth();
    const h = doc.internal.pageSize.getHeight();
    const m = 14;

    doc.setFillColor(...THEME.softBg);
    doc.rect(0, 0, w, h, 'F');

    // Filete superior de 5 colores
    const stripeMargin = 8;
    const stripeY = 8;
    const stripeH = 1.2;
    const stripeW = (w - (stripeMargin * 2)) / 5;

    doc.setFillColor(...THEME.brandRed);
    doc.rect(stripeMargin, stripeY, stripeW, stripeH, 'F');

    doc.setFillColor(...THEME.brandOrange);
    doc.rect(stripeMargin + stripeW, stripeY, stripeW, stripeH, 'F');

    doc.setFillColor(...THEME.brandViolet);
    doc.rect(stripeMargin + (stripeW * 2), stripeY, stripeW, stripeH, 'F');

    doc.setFillColor(...THEME.brandGreen);
    doc.rect(stripeMargin + (stripeW * 3), stripeY, stripeW, stripeH, 'F');

    doc.setFillColor(...THEME.brandBlue);
    doc.rect(stripeMargin + (stripeW * 4), stripeY, stripeW, stripeH, 'F');

    // BLOQUE SUPERIOR CENTRADO
    const topBlockW = w - 40;
    const topBlockX = (w - topBlockW) / 2;

    if (logoDataUrl) {
        const logoBoxW = 52;
        const logoBoxH = 18;
        const logoBoxX = (w - logoBoxW) / 2;
        const logoBoxY = 16;

        drawRoundedCard(doc, logoBoxX, logoBoxY, logoBoxW, logoBoxH, THEME.white, THEME.line, 4);
        drawPdfLogo(doc, logoDataUrl, logoBoxX + 4, logoBoxY + 3, 44);
    }

    // título
    setFont(doc, 'bold', 24, THEME.ink);
    doc.text(title, w / 2, 52, { align: 'center' });

    // subtítulo
    let subtitleBottomY = 52;
    if (subtitle) {
        setFont(doc, 'normal', 11, THEME.muted);
        const subtitleLines = doc.splitTextToSize(subtitle, topBlockW);
        doc.text(subtitleLines, w / 2, 60, { align: 'center' });
        subtitleBottomY = 60 + ((subtitleLines.length - 1) * 5);
    }

    // ficha
    const metaY = subtitle ? subtitleBottomY + 10 : 64;
    const metaW = w - (m * 2);
    const metaH = 28;

    drawRoundedCard(doc, m, metaY, metaW, metaH, THEME.white, THEME.line, 5);

    setFont(doc, 'bold', 10, THEME.ink);
    doc.text('Fitxa de l’informe', m + 4, metaY + 8);

    setFont(doc, 'normal', 9, THEME.text);
    doc.text(`Data de generació: ${generatedAt}`, m + 4, metaY + 15);
    doc.text(`Persones / equips: ${people || '—'}`, m + 4, metaY + 21);

    // TARJETAS VERTICALES
    const cardGap = 8;
    const cardX = m;
    const cardW = w - (m * 2);
    const cardH = 42;

    let cardY = metaY + metaH + 12;

    const panels = [
        {
            color: THEME.brandRed,
            title: 'PARA',
            subtitle: 'ATURA L’IMPULS',
            body: 'Intervenir en calent ens porta a la reacció impulsiva. Parar és regular-nos per poder actuar millor.',
            quote: '“Ara mateix parem aquí... ho reprenem en un moment?”'
        },
        {
            color: THEME.brandViolet,
            title: 'PARLA',
            subtitle: 'DONA VEU REAL',
            body: 'No és només deixar parlar, és escoltar de veritat. Passar de jutjar a entendre.',
            quote: 'Què ha passat segons tu? Com t’ha afectat? Què necessites ara?'
        },
        {
            color: THEME.brandBlue,
            title: 'PACTA',
            subtitle: 'CONSTRUEIX COMPROMÍS',
            body: 'No busquem el càstig, busquem la reparació construïda. Responsabilitat no és culpa: és capacitat d’actuar.',
            quote: 'Què pots fer per reparar el dany? Què necessites per ser reparat?'
        }
    ];

    panels.forEach(panel => {
        doc.setFillColor(...panel.color);
        doc.roundedRect(cardX, cardY, cardW, cardH, 8, 8, 'F');

        // Título
        setFont(doc, 'bold', 18, THEME.white);
        doc.text(panel.title, cardX + 8, cardY + 11);

        // Subtítulo
        setFont(doc, 'bold', 9.5, THEME.white);
        doc.text(panel.subtitle, cardX + 8, cardY + 19);

        // Cuerpo
        setFont(doc, 'normal', 9.5, THEME.white);
        const bodyLines = doc.splitTextToSize(panel.body, cardW - 92);
        doc.text(bodyLines, cardX + 8, cardY + 28);

        // Pastilla derecha
        const quoteBoxW = 64;
        const quoteBoxH = 18;
        const quoteBoxX = cardX + cardW - quoteBoxW - 8;
        const quoteBoxY = cardY + 12;

        doc.setFillColor(30, 30, 32);
        doc.roundedRect(quoteBoxX, quoteBoxY, quoteBoxW, quoteBoxH, 4, 4, 'F');

        setFont(doc, 'bold', 7.5, THEME.white);
        const quoteLines = doc.splitTextToSize(panel.quote, quoteBoxW - 8).slice(0, 3);
        doc.text(quoteLines, quoteBoxX + 4, quoteBoxY + 5);

        cardY += cardH + cardGap;
    });

    // Observaciones
    if (notes) {
        const notesY = cardY + 2;
        const notesH = 18;

        if (notesY + notesH < h - 18) {
            drawRoundedCard(doc, m, notesY, w - (m * 2), notesH, THEME.white, THEME.line, 4);

            setFont(doc, 'bold', 8.5, THEME.ink);
            doc.text('Observacions', m + 4, notesY + 6);

            setFont(doc, 'normal', 8.2, THEME.text);
            const noteLines = doc.splitTextToSize(notes, w - (m * 2) - 8).slice(0, 2);
            doc.text(noteLines, m + 4, notesY + 12);
        }
    }

    drawPageFooter(doc, {
        generatedAt,
        pageNumber: 1,
        pageWidth: w,
        pageHeight: h,
        margin: m
    });
}

function drawMapPage(doc, {
  logoDataUrl,
  title,
  subtitle,
  generatedAt,
  imgData
}) {
  const pageWidth = doc.internal.pageSize.getWidth();
  const pageHeight = doc.internal.pageSize.getHeight();
  const margin = 10;
  const contentWidth = pageWidth - margin * 2;

  const headerBottom = drawPageHeader(doc, {
    logoDataUrl,
    title: title,
    subtitle: subtitle || 'Mapa visual de la sessió',
    pageWidth,
    margin
  });

  doc.setFont('helvetica', 'normal');
  doc.setFontSize(8.5);
  doc.setTextColor(100, 116, 139);
  doc.text(`Generat: ${generatedAt}`, margin + 2, headerBottom + 6);

  const frameY = headerBottom + 10;
  const frameH = pageHeight - frameY - 12;

  doc.setFillColor(255, 255, 255);
  doc.roundedRect(margin, frameY, contentWidth, frameH, 5, 5, 'F');
  doc.setDrawColor(226, 232, 240);
  doc.roundedRect(margin, frameY, contentWidth, frameH, 5, 5, 'S');

  const imgProps = doc.getImageProperties(imgData);
  let imgWidth = contentWidth - 4;
  let imgHeight = (imgProps.height * imgWidth) / imgProps.width;

  if (imgHeight > frameH - 4) {
    imgHeight = frameH - 4;
    imgWidth = (imgProps.width * imgHeight) / imgProps.height;
  }

  const imgX = margin + (contentWidth - imgWidth) / 2;
  const imgY = frameY + (frameH - imgHeight) / 2;

  doc.addImage(imgData, 'PNG', imgX, imgY, imgWidth, imgHeight);

  drawPageFooter(doc, {
    generatedAt,
    pageNumber: 2,
    pageWidth,
    pageHeight,
    margin
  });
}

async function exportOrganizerPdf() {
    const { jsPDF } = window.jspdf;

    const title = sanitizeUserText(document.getElementById('export-title')?.value || '');
    const subtitle = sanitizeUserText(document.getElementById('export-subtitle')?.value || '') || '';
    const people = sanitizeUserText(document.getElementById('export-people')?.value || '') || '';
    const notes = sanitizeUserText(document.getElementById('export-notes')?.value || '') || '';
    const generatedAt = getFormattedDateTime();

    if (!title) {
        alert('Has d’indicar un nom per generar el PDF.');
        document.getElementById('export-title')?.focus();
        return;
    }

    const board = document.querySelector('#view-organitzador .organizer-layout');
    if (!board) return;

    const situationTexts = getBoardTagTexts('situacions-area').map(cleanExportTagText);
    const emotionTexts = getBoardTagTexts('emocions-area-organizer').map(cleanExportTagText);
    const needTexts = getBoardTagTexts('necessitats-area').map(cleanExportTagText);
    const actionTexts = getBoardTagTexts('accions-area').map(cleanExportTagText);

    const customSituationTexts = customSituations.map(cleanExportTagText);
    const customActionTexts = customActions.map(cleanExportTagText);

    const templateName =
        currentTemplate && currentTemplate !== 'blank'
            ? getTemplateLabel(currentTemplate)
            : 'En blanc';

    const executiveSummary = buildExecutiveSummaryData({
        situationTexts,
        emotionTexts,
        needTexts,
        actionTexts,
        customSituationTexts,
        customActionTexts,
        notes,
        context: templateName
    });

    const templateGuidance = {
        situations: templates?.[currentTemplate]?.situacions || '',
        emotions: templates?.[currentTemplate]?.emocions || '',
        needs: templates?.[currentTemplate]?.necessitats || '',
        actions: templates?.[currentTemplate]?.accions || ''
    };

    try {
        let logoDataUrl = null;

        try {
            logoDataUrl = await loadImageAsDataUrl(POLITENICS_LOGO_URL);
        } catch (e) {
            console.warn('No s’ha pogut carregar el logo del Politècnics', e);
        }

        const previousBoxShadow = board.style.boxShadow;
        board.style.boxShadow = 'none';

        const canvas = await html2canvas(board, {
            backgroundColor: '#ffffff',
            scale: 2,
            useCORS: true
        });

        board.style.boxShadow = previousBoxShadow;

        const imgData = canvas.toDataURL('image/png');

        const pdf = new jsPDF({
            orientation: 'portrait',
            unit: 'mm',
            format: 'a4'
        });

        let pageNumber = 1;

        /* =========================
           PÀGINA 1: PORTADA
           ========================= */
        drawCoverPage(pdf, {
            logoDataUrl,
            title,
            subtitle,
            people,
            notes,
            generatedAt
        });

        /* =========================
           PÀGINA 2: MAPA
           ========================= */
        pdf.addPage('a4', 'landscape');
        pageNumber += 1;

        {
            const pageWidth = pdf.internal.pageSize.getWidth();
            const pageHeight = pdf.internal.pageSize.getHeight();
            const margin = 10;
            const contentWidth = pageWidth - (margin * 2);

            const headerBottom = drawPageHeader(pdf, {
                logoDataUrl,
                title,
                subtitle: subtitle || 'Mapa visual de la sessió',
                pageWidth,
                margin,
                color: [255, 255, 255]
            });

            pdf.setFont('helvetica', 'normal');
            pdf.setFontSize(8.5);
            pdf.setTextColor(95, 95, 95);
            pdf.text('Vista completa del mapa exportat', margin + 2, headerBottom + 6);

            const frameY = headerBottom + 10;
            const frameH = pageHeight - frameY - 14;

            pdf.setFillColor(255, 255, 255);
            pdf.setDrawColor(170, 170, 170);
            pdf.roundedRect(margin, frameY, contentWidth, frameH, 5, 5, 'S');

            const imgProps = pdf.getImageProperties(imgData);

            const maxImgWidth = contentWidth - 4;
            const maxImgHeight = frameH - 4;

            let imgWidth = maxImgWidth;
            let imgHeight = (imgProps.height * imgWidth) / imgProps.width;

            if (imgHeight > maxImgHeight) {
                imgHeight = maxImgHeight;
                imgWidth = (imgProps.width * imgHeight) / imgProps.height;
            }

            const imgX = margin + ((contentWidth - imgWidth) / 2);
            const imgY = frameY + ((frameH - imgHeight) / 2);

            pdf.addImage(imgData, 'PNG', imgX, imgY, imgWidth, imgHeight);

            drawPageFooter(pdf, {
                generatedAt,
                pageNumber,
                pageWidth,
                pageHeight,
                margin
            });
        }

        /* =========================
           PÀGINA 3: RESUM EXECUTIU
           ========================= */
        pdf.addPage('a4', 'portrait');
        pageNumber += 1;

        drawExecutiveSummaryPage(pdf, {
            logoDataUrl,
            generatedAt,
            title,
            subtitle,
            people,
            notes,
            summary: executiveSummary,
            pageNumber,
            THEME
        });

        /* =========================
           PÀGINA 4+: DESENVOLUPAMENT DEL REGISTRE
           ========================= */
        pdf.addPage('a4', 'portrait');
        pageNumber += 1;

        pageNumber = drawRegistryContinuationPage(pdf, {
            generatedAt,
            title,
            summary: executiveSummary,
            templateGuidance,
            customSituationTexts,
            customActionTexts,
            situationTexts,
            emotionTexts,
            needTexts,
            actionTexts,
            pageNumber
        });

        const safeFileName = title
            .toLowerCase()
            .normalize('NFD')
            .replace(/[\u0300-\u036f]/g, '')
            .replace(/[^a-z0-9]+/g, '-')
            .replace(/^-+|-+$/g, '')
            || 'resum-mapa';

        pdf.save(`${safeFileName}.pdf`);
        closeExportModal();

    } catch (error) {
        console.error('Error generant el PDF:', error);
        alert('No s’ha pogut generar el PDF.');
    }
}


    function setFont(doc, style = 'normal', size = 10, color = THEME.ink) {
        doc.setFont('helvetica', style);
        doc.setFontSize(size);
        doc.setTextColor(...color);
    }

    function drawRoundedCard(doc, x, y, w, h, fillColor = THEME.white, borderColor = THEME.line, radius = 4) {
        doc.setFillColor(...fillColor);
        doc.roundedRect(x, y, w, h, radius, radius, 'F');
        doc.setDrawColor(...borderColor);
        doc.roundedRect(x, y, w, h, radius, radius, 'S');
    }

    function drawSectionCard(doc, {
        x,
        y,
        w,
        h,
        title,
        fillColor = THEME.white,
        accentColor = THEME.brandRed
    }) {
        drawRoundedCard(doc, x, y, w, h, fillColor, THEME.line, 4);

        doc.setFillColor(...accentColor);
        doc.roundedRect(x, y, 3, h, 2, 2, 'F');

        setFont(doc, 'bold', 11, THEME.ink);
        doc.text(title, x + 7, y + 8);
    }

    function drawStatCard(doc, x, y, w, h, value, label, accentColor) {
        drawRoundedCard(doc, x, y, w, h, THEME.white, THEME.line, 4);

        doc.setFillColor(...accentColor);
        doc.roundedRect(x, y, w, 3, 1.5, 1.5, 'F');

        setFont(doc, 'bold', 16, THEME.ink);
        doc.text(String(value), x + 4, y + 10);

        setFont(doc, 'normal', 8.5, THEME.muted);
        doc.text(label, x + 4, y + 16);
    }

    function drawCardText(doc, x, y, text, maxWidth, lineHeight = 4.8, style = 'normal', size = 9.5, color = THEME.text) {
        if (!text) return y;
        setFont(doc, style, size, color);
        const lines = doc.splitTextToSize(text, maxWidth);
        doc.text(lines, x, y);
        return y + lines.length * lineHeight;
    }

    function drawListInCard(doc, x, y, items, maxWidth, maxItems = 8, bulletColor = THEME.ink, textColor = THEME.text) {
        setFont(doc, 'normal', 9.3, textColor);

        if (!items.length) {
            doc.text('—', x, y);
            return y + 5;
        }

        const visibleItems = items.slice(0, maxItems);
        let cursorY = y;

        visibleItems.forEach(item => {
            const bullet = '• ';
            const lines = doc.splitTextToSize(`${bullet}${item}`, maxWidth);
            doc.setTextColor(...textColor);
            doc.text(lines, x, cursorY);
            cursorY += (lines.length * 4.5) + 1.2;
        });

        if (items.length > maxItems) {
            setFont(doc, 'normal', 8.5, THEME.muted);
            doc.text(`+ ${items.length - maxItems} elements més`, x, cursorY);
            cursorY += 5;
        }

        return cursorY;
    }

    function estimateTextHeight(doc, text, maxWidth, lineHeight = 4.8, fontSize = 9.5) {
        if (!text) return 0;
        doc.setFont('helvetica', 'normal');
        doc.setFontSize(fontSize);
        const lines = doc.splitTextToSize(text, maxWidth);
        return lines.length * lineHeight;
    }

    function estimateListHeight(doc, items, maxWidth, maxItems = 8, lineHeight = 4.5) {
        if (!items.length) return 14;

        doc.setFont('helvetica', 'normal');
        doc.setFontSize(9.3);

        const visibleItems = items.slice(0, maxItems);
        let linesCount = 0;

        visibleItems.forEach(item => {
            const lines = doc.splitTextToSize(`• ${item}`, maxWidth);
            linesCount += lines.length;
        });

        if (items.length > maxItems) linesCount += 1;

        return 16 + (linesCount * lineHeight);
    }

    function drawPdfLogo(doc, logoDataUrl, x, y, targetW = 30) {
    if (!logoDataUrl) {
        console.warn('Logo no disponible: logoDataUrl es null');
        return null;
    }

    try {
        const props = doc.getImageProperties(logoDataUrl);
        const aspectRatio = props.width / props.height;
        const targetH = targetW / aspectRatio;

        doc.addImage(logoDataUrl, 'PNG', x, y, targetW, targetH);
        return { width: targetW, height: targetH };
    } catch (e) {
        console.warn('No s’ha pogut dibuixar el logo al PDF', e);
        return null;
    }
}
    function drawPageHeader(doc, {
        logoDataUrl,
        title,
        subtitle = '',
        pageWidth,
        margin = 14,
        fillColor = THEME.softBg
    }) {
        const headerY = 10;
        const headerH = 24;
        const contentW = pageWidth - (margin * 2);

        // Filete superior de 5 colores Politècnics
        const stripeY = headerY - 2.2;
        const stripeH = 1.1;
        const stripeW = contentW / 5;

        doc.setFillColor(...THEME.brandRed);
        doc.rect(margin, stripeY, stripeW, stripeH, 'F');

        doc.setFillColor(...THEME.brandOrange);
        doc.rect(margin + stripeW, stripeY, stripeW, stripeH, 'F');

        doc.setFillColor(...THEME.brandViolet);
        doc.rect(margin + (stripeW * 2), stripeY, stripeW, stripeH, 'F');

        doc.setFillColor(...THEME.brandGreen);
        doc.rect(margin + (stripeW * 3), stripeY, stripeW, stripeH, 'F');

        doc.setFillColor(...THEME.brandBlue);
        doc.rect(margin + (stripeW * 4), stripeY, stripeW, stripeH, 'F');

        drawRoundedCard(doc, margin, headerY, contentW, headerH, fillColor, THEME.line, 4);

        setFont(doc, 'bold', 15, THEME.ink);
        doc.text(title, margin + 4, headerY + 9);

        if (subtitle) {
            setFont(doc, 'normal', 9, THEME.muted);
            doc.text(subtitle, margin + 4, headerY + 16);
        }

        if (logoDataUrl) {
            const logoBoxW = 42;
            const logoBoxH = 16;
            const logoBoxX = pageWidth - margin - logoBoxW - 2;
            const logoBoxY = headerY + 4;

            drawRoundedCard(doc, logoBoxX, logoBoxY, logoBoxW, logoBoxH, THEME.white, THEME.line, 3);

            const targetLogoW = 30;
            const props = doc.getImageProperties(logoDataUrl);
            const targetLogoH = targetLogoW * (props.height / props.width);

            const logoX = logoBoxX + ((logoBoxW - targetLogoW) / 2);
            const logoY = logoBoxY + ((logoBoxH - targetLogoH) / 2);

            drawPdfLogo(doc, logoDataUrl, logoX, logoY, targetLogoW);
        }
        return headerY + headerH;
}

    
    
    
    function drawPageFooter(doc, {
        generatedAt,
        pageNumber,
        pageWidth,
        pageHeight,
        margin = 14
    }) {
        const y = pageHeight - 8;

        doc.setDrawColor(...THEME.line);
        doc.line(margin, y - 4, pageWidth - margin, y - 4);

        setFont(doc, 'normal', 8, THEME.muted);
        doc.text(`Generat: ${generatedAt}`, margin, y);
        doc.text(`Pàgina ${pageNumber}`, pageWidth - margin, y, { align: 'right' });
    }

function drawCoverPage(doc, {
    logoDataUrl,
    title,
    subtitle,
    people,
    notes,
    generatedAt
}) {
    const w = doc.internal.pageSize.getWidth();
    const h = doc.internal.pageSize.getHeight();
    const m = 8;

    doc.setFillColor(...THEME.softBg);
    doc.rect(0, 0, w, h, 'F');

    // Filete superior de 5 colors
    const stripeMargin = 8;
    const stripeY = 8;
    const stripeH = 1.2;
    const stripeW = (w - (stripeMargin * 2)) / 5;

    doc.setFillColor(...THEME.brandRed);
    doc.rect(stripeMargin, stripeY, stripeW, stripeH, 'F');

    doc.setFillColor(...THEME.brandOrange);
    doc.rect(stripeMargin + stripeW, stripeY, stripeW, stripeH, 'F');

    doc.setFillColor(...THEME.brandViolet);
    doc.rect(stripeMargin + (stripeW * 2), stripeY, stripeW, stripeH, 'F');

    doc.setFillColor(...THEME.brandGreen);
    doc.rect(stripeMargin + (stripeW * 3), stripeY, stripeW, stripeH, 'F');

    doc.setFillColor(...THEME.brandBlue);
    doc.rect(stripeMargin + (stripeW * 4), stripeY, stripeW, stripeH, 'F');

    // Logo fora del bloc agrupat
    if (logoDataUrl) {
        const logoBoxW = 52;
        const logoBoxH = 18;
        const logoBoxX = w - m - logoBoxW;
        const logoBoxY = 16;

        drawRoundedCard(doc, logoBoxX, logoBoxY, logoBoxW, logoBoxH, THEME.white, THEME.line, 4);
        drawPdfLogo(doc, logoDataUrl, logoBoxX + 4, logoBoxY + 3, 44);
    }

    // BLOQUE SUPERIOR AGRUPADO
    const heroX = m;
    const heroY = 40;
    const heroW = w - (m * 2);
    const innerX = heroX + 6;
    const innerW = heroW - 12;

    const subtitleLines = subtitle ? doc.splitTextToSize(subtitle, innerW) : [];
    const subtitleHeight = subtitleLines.length ? subtitleLines.length * 5 + 3 : 0;

    // Observaciones dinámicas
    const notesTitleH = notes ? 6 : 0;
    const notesLines = notes ? doc.splitTextToSize(notes, innerW) : [];
    const maxNotesLines = 8;
    const visibleNotesLines = notesLines.slice(0, maxNotesLines);
    const notesTextH = visibleNotesLines.length ? visibleNotesLines.length * 4.2 : 0;
    const notesBlockH = notes ? 8 + notesTitleH + notesTextH + 4 : 0;

    const heroTopPad = 18;
    const titleBlockH = 9;
    const separator1H = 8;
    const fitxaTitleH = 8;
    const fitxaBodyH = 14;
    const separator2H = notes ? 7 : 0;
    const heroBottomPad = 8;

    const heroH =
        heroTopPad +
        titleBlockH +
        subtitleHeight +
        separator1H +
        fitxaTitleH +
        fitxaBodyH +
        separator2H +
        notesBlockH +
        heroBottomPad;

    drawRoundedCard(doc, heroX, heroY, heroW, heroH, THEME.white, THEME.line, 6);

    let y = heroY + 18;

    // Títol
    setFont(doc, 'bold', 22, THEME.ink);
    doc.text(title, innerX, y);
    y += 9;

    // Subtítol
    if (subtitleLines.length) {
        setFont(doc, 'normal', 11, THEME.muted);
        doc.text(subtitleLines, innerX, y);
        y += subtitleLines.length * 5 + 3;
    }

    // Separador
    doc.setDrawColor(...THEME.line);
    doc.line(innerX, y, heroX + heroW - 6, y);
    y += 8;

    // Fitxa
    setFont(doc, 'bold', 10, THEME.ink);
    doc.text('Fitxa de l’informe', innerX, y);
    y += 8;

    setFont(doc, 'normal', 9.5, THEME.text);
    doc.text(`Data de generació: ${generatedAt}`, innerX, y);
    y += 7;
    doc.text(`Persones / equips: ${people || '—'}`, innerX, y);
    y += 9;

    // Observacions
    if (notes) {
        doc.setDrawColor(...THEME.line);
        doc.line(innerX, y - 2, heroX + heroW - 6, y - 2);

        setFont(doc, 'bold', 9, THEME.ink);
        doc.text('Observacions', innerX, y + 3);

        setFont(doc, 'normal', 8.5, THEME.text);
        doc.text(visibleNotesLines, innerX, y + 9);

        if (notesLines.length > maxNotesLines) {
            const moreY = y + 9 + (visibleNotesLines.length * 4.2) + 1;
            setFont(doc, 'normal', 8, THEME.muted);
            doc.text(`+ ${notesLines.length - maxNotesLines} línies més`, innerX, moreY);
        }
    }

    // TARJETAS VERTICALES AUTOAJUSTADAS
    const cardX = m;
    const cardW = w - (m * 2);
    const cardGap = 4;
    const cardY = heroY + heroH + 8;

    // Reservamos espacio para footer
    const footerSafeTop = h - 18;
    const availableForCards = footerSafeTop - cardY;

    // Altura automática de las 3 tarjetas
    let cardH = (availableForCards - (cardGap * 2)) / 3;

    // límites razonables
    cardH = Math.max(32, Math.min(42, cardH));

    const panels = [
        {
            color: THEME.brandRed,
            title: 'PARA',
            subtitle: 'ATURA L’IMPULS',
            body: 'Intervenir en calent ens porta a la reacció impulsiva. Parar és regular-nos per poder actuar millor.',
            quote: '“Ara mateix parem aquí... ho reprenem en un moment?”'
        },
        {
            color: THEME.brandViolet,
            title: 'PARLA',
            subtitle: 'DONA VEU REAL',
            body: 'No és només deixar parlar, és escoltar de veritat. Passar de jutjar a entendre.',
            quote: 'Què ha passat segons tu? Com t’ha afectat? Què necessites ara?'
        },
        {
            color: THEME.brandBlue,
            title: 'PACTA',
            subtitle: 'CONSTRUEIX COMPROMÍS',
            body: 'No busquem el càstig, busquem la reparació construïda. Responsabilitat no és culpa: és capacitat d’actuar.',
            quote: 'Què pots fer per reparar el dany? Què necessites per ser reparat?'
        }
    ];

    let currentCardY = cardY;

    panels.forEach(panel => {
        doc.setFillColor(...panel.color);
        doc.roundedRect(cardX, currentCardY, cardW, cardH, 8, 8, 'F');

        setFont(doc, 'bold', 18, THEME.white);
        doc.text(panel.title, cardX + 8, currentCardY + 10);

        setFont(doc, 'bold', 9.2, THEME.white);
        doc.text(panel.subtitle, cardX + 8, currentCardY + 18);

        const quoteBoxW = 64;
        const quoteBoxH = Math.min(18, cardH - 10);
        const quoteBoxX = cardX + cardW - quoteBoxW - 8;
        const quoteBoxY = currentCardY + Math.max(6, (cardH - quoteBoxH) / 2);

        const bodyW = cardW - 92;
        const bodyMaxLines = cardH < 36 ? 2 : 3;

        setFont(doc, 'normal', 9, THEME.white);
        const bodyLines = doc.splitTextToSize(panel.body, bodyW).slice(0, bodyMaxLines);
        doc.text(bodyLines, cardX + 8, currentCardY + 26);

        doc.setFillColor(30, 30, 32);
        doc.roundedRect(quoteBoxX, quoteBoxY, quoteBoxW, quoteBoxH, 4, 4, 'F');

        setFont(doc, 'bold', 7.2, THEME.white);
        const quoteMaxLines = cardH < 36 ? 2 : 3;
        const quoteLines = doc.splitTextToSize(panel.quote, quoteBoxW - 8).slice(0, quoteMaxLines);
        doc.text(quoteLines, quoteBoxX + 4, quoteBoxY + 5);

        currentCardY += cardH + cardGap;
    });

    drawPageFooter(doc, {
        generatedAt,
        pageNumber: 1,
        pageWidth: w,
        pageHeight: h,
        margin: m
    });
}

function addWrappedLabelValue(doc, x, y, label, value, maxWidth) {
        const text = `${label}${value || '—'}`;
        return drawCardText(doc, x, y, text, maxWidth, 4.8, 'normal', 9.3, THEME.text);
    }
const THEME_MONO = {
    black: [20, 20, 20],
    text: [45, 45, 45],
    muted: [110, 110, 110],
    line: [190, 190, 190],
    softLine: [225, 225, 225],
    white: [255, 255, 255],
    panel: [248, 248, 248]
};

const THEME_OLD = {
    ink: [20, 20, 20],
    text: [45, 45, 45],
    muted: [95, 95, 95],
    line: [190, 190, 190],
    softBg: [248, 248, 248],
    white: [255, 255, 255],

    brandRed: [60, 60, 60],
    brandOrange: [90, 90, 90],
    brandViolet: [110, 110, 110],
    brandGreen: [130, 130, 130],
    brandBlue: [80, 80, 80],

    situationBg: [252, 252, 252],
    emotionBg: [250, 250, 250],
    needBg: [248, 248, 248],
    actionBg: [246, 246, 246]
};
    const THEME = {
        brandRed: [216, 38, 46],       // #D8262E
        brandOrange: [249, 141, 41],   // #F98D29
        brandViolet: [106, 71, 141],   // #6A478D
        brandGreen: [100, 139, 26],    // #648B1A
        brandBlue: [0, 154, 190],      // #009ABE

        // alias por compatibilidad con código anterior
        //brandPurple: [106, 71, 141],
        //brandTeal: [0, 154, 190],
        //brandAmber: [249, 141, 41],

        ink: [24, 24, 27],
        text: [87, 87, 86],
        muted: [100, 116, 139],
        line: [226, 232, 240],
        softBg: [244, 247, 249],
        white: [255, 255, 255],

        situationBg: [255, 245, 245],
        emotionBg: [248, 243, 255],
        needBg: [240, 253, 244],
        actionBg: [240, 249, 255]
        };


        function drawSimpleSection(doc, { x, y, w, h, title }) {
    doc.setFillColor(255, 255, 255);
    doc.setDrawColor(190, 190, 190);
    doc.rect(x, y, w, h);

    doc.setFont('helvetica', 'bold');
    doc.setFontSize(10);
    doc.setTextColor(20, 20, 20);
    doc.text(title.toUpperCase(), x + 4, y + 7);

    doc.setDrawColor(225, 225, 225);
    doc.line(x + 4, y + 10, x + w - 4, y + 10);
}

function drawMonoHeader(doc, {
    title,
    subtitle = '',
    pageWidth,
    margin = 16
}) {
    const y = 14;

    doc.setFont('helvetica', 'bold');
    doc.setFontSize(15);
    doc.setTextColor(20, 20, 20);
    doc.text(title.toUpperCase(), margin, y);

    if (subtitle) {
        doc.setFont('helvetica', 'normal');
        doc.setFontSize(9);
        doc.setTextColor(90, 90, 90);
        doc.text(subtitle, margin, y + 6);
    }

    doc.setDrawColor(160, 160, 160);
    doc.setLineWidth(0.4);
    doc.line(margin, y + 10, pageWidth - margin, y + 10);

    return y + 14;
}

function drawMonoSectionTitle(doc, text, x, y, w) {
    doc.setFont('helvetica', 'bold');
    doc.setFontSize(10);
    doc.setTextColor(20, 20, 20);
    doc.text(text.toUpperCase(), x, y);

    doc.setDrawColor(210, 210, 210);
    doc.setLineWidth(0.2);
    doc.line(x, y + 2, x + w, y + 2);
}

function drawMonoTable(doc, {
    x,
    y,
    w,
    rows,
    colRatios = [0.72, 0.28],
    header = null,
    rowH = 8
}) {
    const col1 = w * colRatios[0];
    const col2 = w * colRatios[1];
    let cursorY = y;

    const drawRow = (left, right, isHeader = false) => {
        doc.setDrawColor(170, 170, 170);
        doc.setLineWidth(0.2);
        doc.rect(x, cursorY, col1, rowH);
        doc.rect(x + col1, cursorY, col2, rowH);

        doc.setFont('helvetica', isHeader ? 'bold' : 'normal');
        doc.setFontSize(9);
        doc.setTextColor(25, 25, 25);
        doc.text(String(left), x + 3, cursorY + 5.2);
        doc.text(String(right), x + col1 + 3, cursorY + 5.2);

        cursorY += rowH;
    };

    if (header) drawRow(header[0], header[1], true);
    rows.forEach(([left, right]) => drawRow(left, right, false));

    return cursorY;
}

function drawMonoLines(doc, {
    x,
    y,
    w,
    lines,
    lineHeight = 5
}) {
    let cursorY = y;

    doc.setFont('helvetica', 'normal');
    doc.setFontSize(9.2);
    doc.setTextColor(40, 40, 40);

    lines.forEach(line => {
        const wrapped = doc.splitTextToSize(line, w);
        doc.text(wrapped, x, cursorY);
        cursorY += wrapped.length * lineHeight;
    });

    return cursorY;
}

function drawRegistryContinuationPage(doc, {
    generatedAt,
    title,
    summary,
    templateGuidance,
    customSituationTexts,
    customActionTexts,
    situationTexts,
    emotionTexts,
    needTexts,
    actionTexts,
    pageNumber
}) {
    const w = doc.internal.pageSize.getWidth();
    const h = doc.internal.pageSize.getHeight();
    const m = 16;
    const usableW = w - (m * 2);

    let y = drawMonoHeader(doc, {
        title: 'Desarrollo del registro',
        subtitle: title || '',
        pageWidth: w,
        margin: m
    });

    const drawBlock = (sectionTitle, groups) => {
        drawMonoSectionTitle(doc, sectionTitle, m, y, usableW);
        y += 8;

        groups.forEach(group => {
            const hasItems = Array.isArray(group.items) && group.items.length > 0;
            const guidance = group.guidance || '';

            doc.setFont('helvetica', 'bold');
            doc.setFontSize(9.2);
            doc.setTextColor(20, 20, 20);
            doc.text(group.label, m, y);
            y += 5;

            if (guidance) {
                doc.setFont('helvetica', 'normal');
                doc.setFontSize(8.6);
                doc.setTextColor(95, 95, 95);
                const guidanceLines = doc.splitTextToSize(`Orientación de plantilla: ${guidance}`, usableW - 2);
                doc.text(guidanceLines, m, y);
                y += guidanceLines.length * 4.3;
            }

            doc.setDrawColor(170, 170, 170);
            doc.setLineWidth(0.2);

            const rows = hasItems
                ? group.items.map(item => [item])
                : [['No consta registro.']];

            rows.forEach(([item]) => {
                const textLines = doc.splitTextToSize(item, usableW - 8);
                const rowH = Math.max(8, (textLines.length * 4.3) + 3);

                if (y + rowH > h - 18) {
                    doc.addPage('a4', 'portrait');
                    pageNumber += 1;
                    y = drawMonoHeader(doc, {
                        title: 'Desenvolupament del registre',
                        subtitle: title || '',
                        pageWidth: w,
                        margin: m
                    });
                }

                doc.rect(m, y - 3.5, usableW, rowH);
                doc.setFont('helvetica', 'normal');
                doc.setFontSize(9);
                doc.setTextColor(40, 40, 40);
                doc.text(textLines, m + 3, y + 1.5);
                y += rowH + 2;
            });

            y += 4;
        });
    };

    drawBlock('1. Situacions', [
        {
            label: 'Situacions donades d’alta en el banc',
            items: customSituationTexts,
            guidance: templateGuidance.situations
        },
        {
            label: 'Situacions seleccionades al mapa',
            items: situationTexts,
            guidance: templateGuidance.situations
        }
    ]);

    if (y > h - 70) {
        doc.addPage('a4', 'portrait');
        pageNumber += 1;
        y = drawMonoHeader(doc, {
            title: 'Desenvolupament del registre',
            subtitle: title || '',
            pageWidth: w,
            margin: m
        });
    }

    drawBlock('2. Emocions', [
        {
            label: 'Emocions registrades',
            items: emotionTexts,
            guidance: templateGuidance.emotions
        }
    ]);

    if (y > h - 70) {
        doc.addPage('a4', 'portrait');
        pageNumber += 1;
        y = drawMonoHeader(doc, {
            title: 'Desenvolupament del registre',
            subtitle: title || '',
            pageWidth: w,
            margin: m
        });
    }

    drawBlock('3. Necessitats', [
        {
            label: 'Necessitats registrades',
            items: needTexts,
            guidance: templateGuidance.needs
        }
    ]);

    if (y > h - 70) {
        doc.addPage('a4', 'portrait');
        pageNumber += 1;
        y = drawMonoHeader(doc, {
            title: 'Desenvolupament del registre',
            subtitle: title || '',
            pageWidth: w,
            margin: m
        });
    }

    drawBlock('4. Acciones', [
        {
            label: 'Accions donades d’alta en el banc',
            items: customActionTexts,
            guidance: templateGuidance.actions
        },
        {
            label: 'Accions seleccionades al mapa',
            items: actionTexts,
            guidance: templateGuidance.actions
        }
    ]);

    drawPageFooter(doc, {
        generatedAt,
        pageNumber,
        pageWidth: w,
        pageHeight: h,
        margin: m
    });

    return pageNumber;
}
