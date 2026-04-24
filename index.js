// index.js - VERSIÓN COMPLETA CON PRIMARY Y SECONDARY KEYWORDS
const express = require('express');
const ExcelJS = require('exceljs');
const cors = require('cors');

const app = express();
app.use(cors());
app.use(express.json({ limit: '50mb' }));

const PORT = process.env.PORT || 8080;

// ============================================
// ESTILOS
// ============================================
const styles = {
  pageTitle: {
    font: { bold: true, size: 14, color: { argb: 'FFFFFFFF' } },
    fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1F4E78' } },
    alignment: { horizontal: 'left', vertical: 'center' }
  },
  sectionHeader: {
    font: { bold: true, size: 11 },
    fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE7E6E6' } },
    alignment: { horizontal: 'left', vertical: 'center' }
  },
  tableHeader: {
    font: { bold: true, size: 11, color: { argb: 'FFFFFFFF' } },
    fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF4472C4' } },
    alignment: { horizontal: 'center', vertical: 'center' },
    border: {
      top: { style: 'thin', color: { argb: 'FF000000' } },
      left: { style: 'thin', color: { argb: 'FF000000' } },
      bottom: { style: 'thin', color: { argb: 'FF000000' } },
      right: { style: 'thin', color: { argb: 'FF000000' } }
    }
  },
  tableCell: {
    alignment: { horizontal: 'left', vertical: 'center' },
    border: {
      top: { style: 'thin', color: { argb: 'FFD3D3D3' } },
      left: { style: 'thin', color: { argb: 'FFD3D3D3' } },
      bottom: { style: 'thin', color: { argb: 'FFD3D3D3' } },
      right: { style: 'thin', color: { argb: 'FFD3D3D3' } }
    }
  },
  tableCellCenter: {
    alignment: { horizontal: 'center', vertical: 'center' },
    border: {
      top: { style: 'thin', color: { argb: 'FFD3D3D3' } },
      left: { style: 'thin', color: { argb: 'FFD3D3D3' } },
      bottom: { style: 'thin', color: { argb: 'FFD3D3D3' } },
      right: { style: 'thin', color: { argb: 'FFD3D3D3' } }
    }
  },
  priorityHigh: {
    font: { bold: true, color: { argb: 'FFFFFFFF' } },
    fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE74C3C' } },
    alignment: { horizontal: 'center', vertical: 'center' },
    border: {
      top: { style: 'thin', color: { argb: 'FFD3D3D3' } },
      left: { style: 'thin', color: { argb: 'FFD3D3D3' } },
      bottom: { style: 'thin', color: { argb: 'FFD3D3D3' } },
      right: { style: 'thin', color: { argb: 'FFD3D3D3' } }
    }
  },
  priorityMedium: {
    font: { bold: true },
    fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF39C12' } },
    alignment: { horizontal: 'center', vertical: 'center' },
    border: {
      top: { style: 'thin', color: { argb: 'FFD3D3D3' } },
      left: { style: 'thin', color: { argb: 'FFD3D3D3' } },
      bottom: { style: 'thin', color: { argb: 'FFD3D3D3' } },
      right: { style: 'thin', color: { argb: 'FFD3D3D3' } }
    }
  },
  priorityLow: {
    font: { bold: true },
    fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF27AE60' } },
    alignment: { horizontal: 'center', vertical: 'center' },
    border: {
      top: { style: 'thin', color: { argb: 'FFD3D3D3' } },
      left: { style: 'thin', color: { argb: 'FFD3D3D3' } },
      bottom: { style: 'thin', color: { argb: 'FFD3D3D3' } },
      right: { style: 'thin', color: { argb: 'FFD3D3D3' } }
    }
  },
  dimensionHeader: {
    font: { bold: true, size: 10, color: { argb: 'FFFFFFFF' } },
    fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF8E44AD' } },
    alignment: { horizontal: 'center', vertical: 'center' },
    border: {
      top: { style: 'thin', color: { argb: 'FF000000' } },
      left: { style: 'thin', color: { argb: 'FF000000' } },
      bottom: { style: 'thin', color: { argb: 'FF000000' } },
      right: { style: 'thin', color: { argb: 'FF000000' } }
    }
  },
  faqHeader: {
    font: { bold: true, size: 11 },
    fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFC000' } },
    alignment: { horizontal: 'left', vertical: 'center' }
  },
  metadataLabel: {
    font: { bold: true },
    alignment: { horizontal: 'left', vertical: 'center' }
  }
};

// ============================================
// HELPER: Generar slug de URL
// ============================================
function generateSlug(clusterName, service, city) {
  const cleanName = (clusterName || service || 'service')
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[^a-z0-9\s-]/g, '')
    .trim()
    .replace(/\s+/g, '-');
  
  const cleanCity = (city || '')
    .toLowerCase()
    .split(',')[0]
    .trim()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[^a-z0-9\s-]/g, '')
    .replace(/\s+/g, '-');

  return cleanCity ? `/${cleanName}-${cleanCity}` : `/${cleanName}`;
}

// ============================================
// HELPER: Aplicar estilo de prioridad
// ============================================
function getPriorityStyle(priority) {
  if (priority === 'HIGH') return styles.priorityHigh;
  if (priority === 'MEDIUM') return styles.priorityMedium;
  if (priority === 'LOW') return styles.priorityLow;
  return styles.tableCellCenter;
}

// ============================================
// ENDPOINT: /generate-excel
// ============================================
app.post('/generate-excel', async (req, res) => {
  try {
    const { clusters, clustering_summary } = req.body;

    if (!clusters || !Array.isArray(clusters)) {
      return res.status(400).json({ error: 'Invalid input: clusters array required' });
    }

    console.log(`📊 Generating Excel with ${clusters.length} clusters...`);

    const workbook = new ExcelJS.Workbook();
    const allClusters = clusters;
    const summary = clustering_summary || {};

    // ============================================
    // SHEET 0: TABLA OVERVIEW
    // ============================================
    function createOverviewTableSheet() {
      const sheet = workbook.addWorksheet('COMPLETE SERVICE PAGES');
      let row = 1;

      sheet.getCell(`A${row}`).value = `COMPLETE ${allClusters.length} CITY SERVICE PAGES`;
      Object.assign(sheet.getCell(`A${row}`), styles.pageTitle);
      sheet.mergeCells(`A${row}:F${row}`);
      sheet.getRow(row).height = 25;
      row += 2;

      const headers = ['#', 'Service', 'City', 'URL', 'Primary Keyword', 'Volume'];
      headers.forEach((header, i) => {
        const cell = sheet.getCell(row, i + 1);
        cell.value = header;
        Object.assign(cell, styles.tableHeader);
      });
      row++;

      allClusters.forEach((cluster, idx) => {
        const service = cluster.primary_dimensions?.service_category || 'N/A';
        const city = (cluster.primary_dimensions?.geographic_scope || '').split(',')[0].trim() || 'N/A';
        const url = generateSlug(cluster.cluster_name, service, city);
        const primaryKeyword = cluster.primary_keyword?.keyword || cluster.keywords?.[0]?.keyword || 'N/A';
        const volume = cluster.primary_keyword?.volume || cluster.keywords?.[0]?.volume || 'TBD';

        [
          { value: idx + 1, style: styles.tableCellCenter },
          { value: service, style: styles.tableCell },
          { value: city, style: styles.tableCell },
          { value: url, style: styles.tableCell },
          { value: primaryKeyword, style: styles.tableCell },
          { value: volume, style: styles.tableCellCenter }
        ].forEach((item, colIdx) => {
          const cell = sheet.getCell(row, colIdx + 1);
          cell.value = item.value;
          Object.assign(cell, item.style);
        });
        row++;
      });

      sheet.getColumn(1).width = 5;
      sheet.getColumn(2).width = 40;
      sheet.getColumn(3).width = 20;
      sheet.getColumn(4).width = 50;
      sheet.getColumn(5).width = 35;
      sheet.getColumn(6).width = 10;
    }

    // ============================================
    // SHEET 1: ARCHITECTURE
    // ============================================
    function createArchitectureSheet() {
      const sheet = workbook.addWorksheet('1. Site Architecture');
      let row = 1;

      sheet.getCell(`A${row}`).value = 'SITE ARCHITECTURE BLUEPRINT';
      Object.assign(sheet.getCell(`A${row}`), styles.pageTitle);
      sheet.mergeCells(`A${row}:F${row}`);
      row += 2;

      ['#', 'Cluster Name', 'Service', 'Location', 'Priority', 'Page Type'].forEach((header, i) => {
        const cell = sheet.getCell(row, i + 1);
        cell.value = header;
        Object.assign(cell, styles.tableHeader);
      });
      row++;

      allClusters.forEach((cluster, idx) => {
        const priority = cluster.seo_strategy?.priority || 'MEDIUM';
        
        sheet.getCell(`A${row}`).value = idx + 1;
        Object.assign(sheet.getCell(`A${row}`), styles.tableCellCenter);
        
        sheet.getCell(`B${row}`).value = cluster.cluster_name || 'N/A';
        Object.assign(sheet.getCell(`B${row}`), styles.tableCell);
        
        sheet.getCell(`C${row}`).value = cluster.primary_dimensions?.service_category || 'N/A';
        Object.assign(sheet.getCell(`C${row}`), styles.tableCell);
        
        sheet.getCell(`D${row}`).value = cluster.primary_dimensions?.geographic_scope || 'N/A';
        Object.assign(sheet.getCell(`D${row}`), styles.tableCell);
        
        sheet.getCell(`E${row}`).value = priority;
        Object.assign(sheet.getCell(`E${row}`), getPriorityStyle(priority));
        
        sheet.getCell(`F${row}`).value = cluster.seo_strategy?.recommended_page_type || 'Service Page';
        Object.assign(sheet.getCell(`F${row}`), styles.tableCell);
        
        row++;
      });
      row++;

      sheet.getCell(`A${row}`).value = 'ARCHITECTURE SUMMARY';
      Object.assign(sheet.getCell(`A${row}`), styles.sectionHeader);
      sheet.mergeCells(`A${row}:B${row}`);
      row++;

      sheet.getCell(`A${row}`).value = 'Total Pages:';
      Object.assign(sheet.getCell(`A${row}`), styles.metadataLabel);
      sheet.getCell(`B${row}`).value = allClusters.length;
      row++;

      const highPriority = allClusters.filter(c => c.seo_strategy?.priority === 'HIGH').length;
      sheet.getCell(`A${row}`).value = 'High Priority Pages:';
      Object.assign(sheet.getCell(`A${row}`), styles.metadataLabel);
      sheet.getCell(`B${row}`).value = highPriority;
      row++;

      sheet.getColumn(1).width = 8;
      sheet.getColumn(2).width = 50;
      sheet.getColumn(3).width = 30;
      sheet.getColumn(4).width = 20;
      sheet.getColumn(5).width = 12;
      sheet.getColumn(6).width = 25;
    }

    // ============================================
    // SHEETS 2-N: CLUSTER PAGES (MEJORADO)
    // ============================================
    function createClusterPageSheet(cluster, index) {
      const pageName = cluster.cluster_name || `Cluster ${index}`;
      const sheetName = `${index}. ${pageName.substring(0, 25)}`;
      const sheet = workbook.addWorksheet(sheetName);

      sheet.getColumn(1).width = 35;
      sheet.getColumn(2).width = 85;

      let row = 1;

      // Título
      sheet.getCell(`A${row}`).value = pageName.toUpperCase();
      Object.assign(sheet.getCell(`A${row}`), styles.pageTitle);
      sheet.mergeCells(`A${row}:B${row}`);
      row += 2;

      // Metadata básica
      [
        ['Cluster ID', cluster.cluster_id || index],
        ['Service Category', cluster.primary_dimensions?.service_category || 'N/A'],
        ['Geographic Scope', cluster.primary_dimensions?.geographic_scope || 'N/A'],
        ['Priority', cluster.seo_strategy?.priority || 'MEDIUM'],
        ['Page Type', cluster.seo_strategy?.recommended_page_type || 'Service Page'],
        ['Business Value', cluster.primary_dimensions?.business_value_tier || 'MEDIUM']
      ].forEach(([label, value]) => {
        sheet.getCell(`A${row}`).value = label;
        Object.assign(sheet.getCell(`A${row}`), styles.metadataLabel);
        sheet.getCell(`B${row}`).value = value;
        row++;
      });
      row++;

      // ============================================
      // NUEVA SECCIÓN: 7 DIMENSIONS JUSTIFICATION
      // ============================================
      const dimensionsHeader = sheet.getCell(`A${row}`);
      dimensionsHeader.value = '7 DIMENSIONS - CLUSTERING JUSTIFICATION';
      Object.assign(dimensionsHeader, styles.sectionHeader);
      sheet.mergeCells(`A${row}:B${row}`);
      row++;

      const dimensions = [
        ['Service Category', cluster.primary_dimensions?.service_category || 'N/A'],
        ['Geographic Scope', cluster.primary_dimensions?.geographic_scope || 'N/A'],
        ['SERP Semantics', cluster.primary_dimensions?.serp_semantics || 'N/A'],
        ['Deep Search Intent', cluster.primary_dimensions?.deep_search_intent || 'N/A'],
        ['Pain Points Cluster', cluster.primary_dimensions?.pain_points_cluster || 'N/A'],
        ['Buyer Persona', cluster.primary_dimensions?.buyer_persona || 'N/A'],
        ['Business Value Tier', cluster.primary_dimensions?.business_value_tier || 'N/A']
      ];

      dimensions.forEach(([dimension, value]) => {
        const dimCell = sheet.getCell(`A${row}`);
        dimCell.value = dimension;
        Object.assign(dimCell, styles.dimensionHeader);
        
        const valCell = sheet.getCell(`B${row}`);
        valCell.value = value;
        Object.assign(valCell, styles.tableCell);
        
        row++;
      });
      row++;

      // ============================================
      // PRIMARY KEYWORD
      // ============================================
      const primaryKwHeader = sheet.getCell(`A${row}`);
      primaryKwHeader.value = 'PRIMARY KEYWORD';
      Object.assign(primaryKwHeader, styles.sectionHeader);
      sheet.mergeCells(`A${row}:B${row}`);
      row++;

      // Table headers
      ['Keyword', 'Volume | Search Intent | Pain Point'].forEach((header, i) => {
        const cell = sheet.getCell(row, i + 1);
        cell.value = header;
        Object.assign(cell, styles.tableHeader);
      });
      row++;

      // Primary keyword data
      if (cluster.primary_keyword) {
        const pk = cluster.primary_keyword;
        
        const kwCell = sheet.getCell(`A${row}`);
        kwCell.value = pk.keyword || 'N/A';
        Object.assign(kwCell, styles.tableCell);
        
        const detailsCell = sheet.getCell(`B${row}`);
        detailsCell.value = `Vol: ${pk.volume || 0} | Intent: ${pk.search_intent || 'N/A'} | Pain: ${pk.pain_point || 'N/A'}`;
        Object.assign(detailsCell, styles.tableCell);
        
        row++;
      } else {
        sheet.getCell(`A${row}`).value = '(No primary keyword)';
        sheet.mergeCells(`A${row}:B${row}`);
        row++;
      }
      row++;

      // ============================================
      // SECONDARY KEYWORDS
      // ============================================
      const secondaryKwHeader = sheet.getCell(`A${row}`);
      secondaryKwHeader.value = 'SECONDARY KEYWORDS';
      Object.assign(secondaryKwHeader, styles.sectionHeader);
      sheet.mergeCells(`A${row}:B${row}`);
      row++;

      if (cluster.secondary_keywords && cluster.secondary_keywords.length > 0) {
        cluster.secondary_keywords.forEach((sk, index) => {
          const kwText = `${index + 1}. ${sk.keyword} (Vol: ${sk.volume || 0}, CPC: $${sk.cpc || 0}, Intent: ${sk.search_intent || 'N/A'})`;
          sheet.getCell(`A${row}`).value = kwText;
          sheet.mergeCells(`A${row}:B${row}`);
          Object.assign(sheet.getCell(`A${row}`), styles.tableCell);
          row++;
        });
      } else {
        sheet.getCell(`A${row}`).value = '(No secondary keywords)';
        sheet.mergeCells(`A${row}:B${row}`);
        row++;
      }
      row++;

      // Content Angle
      if (cluster.content_strategy?.content_angle) {
        const contentAngleHeader = sheet.getCell(`A${row}`);
        contentAngleHeader.value = 'CONTENT ANGLE';
        Object.assign(contentAngleHeader, styles.sectionHeader);
        sheet.mergeCells(`A${row}:B${row}`);
        row++;

        sheet.getCell(`A${row}`).value = cluster.content_strategy.content_angle;
        sheet.mergeCells(`A${row}:B${row}`);
        row += 2;
      }

      // Trust Elements
      if (cluster.trust_elements && cluster.trust_elements.length > 0) {
        const trustHeader = sheet.getCell(`A${row}`);
        trustHeader.value = 'TRUST ELEMENTS RECOMMENDED';
        Object.assign(trustHeader, styles.sectionHeader);
        sheet.mergeCells(`A${row}:B${row}`);
        row++;

        cluster.trust_elements.forEach(element => {
          sheet.getCell(`A${row}`).value = `• ${element}`;
          sheet.mergeCells(`A${row}:B${row}`);
          row++;
        });
        row++;
      }

      // FAQs
      const faqHeader = sheet.getCell(`A${row}`);
      faqHeader.value = 'FAQs (5-6 QUESTIONS PER PAGE)';
      Object.assign(faqHeader, styles.faqHeader);
      sheet.mergeCells(`A${row}:B${row}`);
      row++;

      if (cluster.faqs && cluster.faqs.length > 0) {
        cluster.faqs.forEach((faq, i) => {
          const faqCell = sheet.getCell(`A${row}`);
          faqCell.value = `Q${i + 1}`;
          faqCell.font = { bold: true };
          sheet.getCell(`B${row}`).value = faq;
          row++;
        });
      } else {
        sheet.getCell(`A${row}`).value = '(No FAQs generated)';
        row++;
      }
      row++;

      // Extended Keywords
      if (cluster.extended_keywords && cluster.extended_keywords.length > 0) {
        const extendedHeader = sheet.getCell(`A${row}`);
        extendedHeader.value = 'EXTENDED KEYWORDS (Semantic Depth)';
        Object.assign(extendedHeader, styles.sectionHeader);
        sheet.mergeCells(`A${row}:B${row}`);
        row++;

        cluster.extended_keywords.forEach(kw => {
          sheet.getCell(`A${row}`).value = `• ${kw}`;
          sheet.mergeCells(`A${row}:B${row}`);
          row++;
        });
        row++;
      }

      // USP
      if (cluster.usp_differentiation) {
        const uspHeader = sheet.getCell(`A${row}`);
        uspHeader.value = 'USP / DIFFERENTIATION';
        Object.assign(uspHeader, styles.sectionHeader);
        sheet.mergeCells(`A${row}:B${row}`);
        row++;

        sheet.getCell(`A${row}`).value = cluster.usp_differentiation;
        sheet.mergeCells(`A${row}:B${row}`);
        row++;
      }
    }

    // ============================================
    // GENERAR TODAS LAS HOJAS
    // ============================================
    createOverviewTableSheet();
    createArchitectureSheet();
    allClusters.forEach((cluster, index) => {
      createClusterPageSheet(cluster, index + 2);
    });

    // ============================================
    // GENERAR BUFFER
    // ============================================
    const buffer = await workbook.xlsx.writeBuffer();
    const today = new Date().toISOString().split('T')[0];
    const filename = `${today}-keyword-clusters-complete.xlsx`;

    console.log(`✅ Excel generated: ${filename}`);

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
    res.send(buffer);

  } catch (error) {
    console.error('Error generating Excel:', error);
    res.status(500).json({ error: error.message });
  }
});

// ============================================
// HEALTH CHECK
// ============================================
app.get('/', (req, res) => {
  res.json({ status: 'API running', endpoint: '/generate-excel' });
});

app.listen(PORT, () => {
  console.log(`✅ Excel Generator API running on port ${PORT}`);
});
