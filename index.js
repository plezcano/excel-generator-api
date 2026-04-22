// index.js
const express = require('express');
const ExcelJS = require('exceljs');
const cors = require('cors');

const app = express();
app.use(cors());
app.use(express.json({ limit: '50mb' }));

const PORT = process.env.PORT || 3000;

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
    alignment: { horizontal: 'center', vertical: 'center' }
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
    // SHEET 0: OVERVIEW
    // ============================================
    function createOverviewSheet() {
      const sheet = workbook.addWorksheet('0. Site Overview');
      let row = 1;

      sheet.getCell(`A${row}`).value = 'SITE OVERVIEW & CLUSTERING SUMMARY';
      Object.assign(sheet.getCell(`A${row}`), styles.pageTitle);
      sheet.mergeCells(`A${row}:F${row}`);
      row += 2;

      sheet.getCell(`A${row}`).value = 'CLUSTERING SUMMARY';
      Object.assign(sheet.getCell(`A${row}`), styles.sectionHeader);
      sheet.mergeCells(`A${row}:C${row}`);
      row++;

      [
        ['Total Keywords Analyzed', summary.total_keywords || 0],
        ['Total Clusters Created', summary.total_clusters || allClusters.length],
        ['Avg Keywords per Cluster', summary.avg_keywords_per_cluster || 0]
      ].forEach(([label, value]) => {
        sheet.getCell(`A${row}`).value = label;
        Object.assign(sheet.getCell(`A${row}`), styles.metadataLabel);
        sheet.getCell(`B${row}`).value = value;
        row++;
      });
      row++;

      sheet.getCell(`A${row}`).value = 'CLUSTER DISTRIBUTION';
      Object.assign(sheet.getCell(`A${row}`), styles.sectionHeader);
      sheet.mergeCells(`A${row}:C${row}`);
      row++;

      if (summary.cluster_distribution) {
        [
          ['1 keyword clusters', summary.cluster_distribution['1_keyword'] || 0],
          ['2-5 keywords clusters', summary.cluster_distribution['2-5_keywords'] || 0],
          ['6-10 keywords clusters', summary.cluster_distribution['6-10_keywords'] || 0],
          ['11+ keywords clusters', summary.cluster_distribution['11+_keywords'] || 0]
        ].forEach(([label, value]) => {
          sheet.getCell(`A${row}`).value = label;
          sheet.getCell(`B${row}`).value = value;
          row++;
        });
      }
      row++;

      sheet.getCell(`A${row}`).value = 'TOP CLUSTERS BY KEYWORD COUNT';
      Object.assign(sheet.getCell(`A${row}`), styles.sectionHeader);
      sheet.mergeCells(`A${row}:F${row}`);
      row++;

      ['Rank', 'Cluster Name', 'Keywords', 'Priority', 'Service', 'Location'].forEach((header, i) => {
        const cell = sheet.getCell(row, i + 1);
        cell.value = header;
        Object.assign(cell, styles.tableHeader);
      });
      row++;

      const sortedClusters = [...allClusters].sort((a, b) =>
        (b.keywords?.length || 0) - (a.keywords?.length || 0)
      ).slice(0, 10);

      sortedClusters.forEach((cluster, idx) => {
        sheet.getCell(`A${row}`).value = idx + 1;
        sheet.getCell(`B${row}`).value = cluster.cluster_name || 'N/A';
        sheet.getCell(`C${row}`).value = cluster.keywords?.length || 0;
        sheet.getCell(`D${row}`).value = cluster.seo_strategy?.priority || 'MEDIUM';
        sheet.getCell(`E${row}`).value = cluster.primary_dimensions?.service_category || 'N/A';
        sheet.getCell(`F${row}`).value = cluster.primary_dimensions?.geographic_scope || 'N/A';
        row++;
      });

      sheet.getColumn(1).width = 8;
      sheet.getColumn(2).width = 50;
      sheet.getColumn(3).width = 12;
      sheet.getColumn(4).width = 12;
      sheet.getColumn(5).width = 30;
      sheet.getColumn(6).width = 20;
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
        sheet.getCell(`A${row}`).value = idx + 1;
        sheet.getCell(`B${row}`).value = cluster.cluster_name || 'N/A';
        sheet.getCell(`C${row}`).value = cluster.primary_dimensions?.service_category || 'N/A';
        sheet.getCell(`D${row}`).value = cluster.primary_dimensions?.geographic_scope || 'N/A';
        sheet.getCell(`E${row}`).value = cluster.seo_strategy?.priority || 'MEDIUM';
        sheet.getCell(`F${row}`).value = cluster.seo_strategy?.recommended_page_type || 'Service Page';
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
    // SHEETS 2-N: CLUSTER PAGES
    // ============================================
    function createClusterPageSheet(cluster, index) {
      const pageName = cluster.cluster_name || `Cluster ${index}`;
      const sheetName = `${index}. ${pageName.substring(0, 25)}`;
      const sheet = workbook.addWorksheet(sheetName);

      sheet.getColumn(1).width = 35;
      sheet.getColumn(2).width = 85;

      let row = 1;

      sheet.getCell(`A${row}`).value = pageName.toUpperCase();
      Object.assign(sheet.getCell(`A${row}`), styles.pageTitle);
      sheet.mergeCells(`A${row}:B${row}`);
      row += 2;

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

      const primaryHeader = sheet.getCell(`A${row}`);
      primaryHeader.value = 'PRIMARY KEYWORDS';
      Object.assign(primaryHeader, styles.sectionHeader);
      sheet.mergeCells(`A${row}:B${row}`);
      row++;

      if (cluster.keywords && cluster.keywords.length > 0) {
        cluster.keywords.forEach(kw => {
          sheet.getCell(`A${row}`).value = kw.keyword || 'N/A';
          sheet.getCell(`B${row}`).value = `Vol: ${kw.volume || 0} | Intent: ${kw.search_intent || 'N/A'}`;
          row++;
        });
      } else {
        sheet.getCell(`A${row}`).value = '(No keywords)';
        row++;
      }
      row++;

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

      if (cluster.extended_keywords && cluster.extended_keywords.length > 0) {
        const extendedHeader = sheet.getCell(`A${row}`);
        extendedHeader.value = 'EXTENDED KEYWORDS (Semantic Depth)';
        Object.assign(extendedHeader, styles.sectionHeader);
        sheet.mergeCells(`A${row}:B${row}`);
        row++;

        cluster.extended_keywords.forEach(kw => {
          sheet.getCell(`A${row}`).value = kw;
          row++;
        });
        row++;
      }

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
    createOverviewSheet();
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
