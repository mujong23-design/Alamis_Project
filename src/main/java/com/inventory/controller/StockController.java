package com.inventory.controller;

import com.inventory.domain.Inbound;
import com.inventory.service.InboundService;
import jakarta.servlet.http.HttpServletResponse;
import lombok.RequiredArgsConstructor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;

import java.io.IOException;
import java.net.URLEncoder;
import java.nio.charset.StandardCharsets;
import java.time.LocalDate;
import java.util.Comparator;
import java.util.List;
import java.util.stream.Collectors;

/**
 * 현재 재고 조회 — inbound 테이블 기반.
 * 현재재고 = package_count - outbound_qty
 */
@Controller
@RequestMapping("/stock")
@RequiredArgsConstructor
public class StockController {

    private final InboundService inboundService;

    @GetMapping("/list")
    public String list(@RequestParam(required = false) Boolean onlyAvailable,
                       Model model) {
        model.addAttribute("stocks",        loadStocks(onlyAvailable));
        model.addAttribute("onlyAvailable", Boolean.TRUE.equals(onlyAvailable));
        return "stock/list";
    }

    /** 엑셀(.xlsx) 다운로드 — 현재 화면과 동일한 필터·정렬 적용 */
    @GetMapping("/excel")
    public void downloadExcel(@RequestParam(required = false) Boolean onlyAvailable,
                              HttpServletResponse response) throws IOException {
        List<Inbound> stocks = loadStocks(onlyAvailable);

        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("현재재고");

            // ===== 스타일 정의 =====
            CellStyle headerStyle = workbook.createCellStyle();
            Font headerFont = workbook.createFont();
            headerFont.setBold(true);
            headerFont.setColor(IndexedColors.WHITE.getIndex());
            headerStyle.setFont(headerFont);
            headerStyle.setFillForegroundColor(IndexedColors.GREY_50_PERCENT.getIndex());
            headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            headerStyle.setAlignment(HorizontalAlignment.CENTER);
            headerStyle.setVerticalAlignment(VerticalAlignment.CENTER);
            setBorder(headerStyle);

            CellStyle subHeaderStyle = workbook.createCellStyle();
            subHeaderStyle.cloneStyleFrom(headerStyle);
            Font subFont = workbook.createFont();
            subFont.setColor(IndexedColors.WHITE.getIndex());
            subFont.setItalic(true);
            subHeaderStyle.setFont(subFont);

            CellStyle textCenterStyle = workbook.createCellStyle();
            textCenterStyle.setAlignment(HorizontalAlignment.CENTER);
            setBorder(textCenterStyle);

            CellStyle numStyle = workbook.createCellStyle();
            numStyle.setAlignment(HorizontalAlignment.RIGHT);
            numStyle.setDataFormat(workbook.createDataFormat().getFormat("#,##0"));
            setBorder(numStyle);

            CellStyle decimalStyle = workbook.createCellStyle();
            decimalStyle.setAlignment(HorizontalAlignment.RIGHT);
            decimalStyle.setDataFormat(workbook.createDataFormat().getFormat("0.000"));
            setBorder(decimalStyle);

            // ===== 헤더 (2단) =====
            String[] titles = {"No", "입고번호", "크기", "초기재고", "총입고량", "총출고량", "현재재고"};
            String[] units  = {"",   "",        "(μm)",  "(포장)",  "(포장)",  "(포장)",  "(포장)"};

            Row titleRow = sheet.createRow(0);
            Row unitRow  = sheet.createRow(1);
            for (int i = 0; i < titles.length; i++) {
                Cell t = titleRow.createCell(i);
                t.setCellValue(titles[i]);
                t.setCellStyle(headerStyle);
                Cell u = unitRow.createCell(i);
                u.setCellValue(units[i]);
                u.setCellStyle(subHeaderStyle);
            }

            // ===== 데이터 =====
            int rowNum = 2;
            int no = 1;
            for (Inbound s : stocks) {
                Row row = sheet.createRow(rowNum++);

                cell(row, 0, no++, textCenterStyle);
                cell(row, 1, s.getInboundId(), textCenterStyle);
                cellNumber(row, 2, s.getSizeUm() != null ? s.getSizeUm().doubleValue() : null, decimalStyle);
                cellNumber(row, 3, s.getPackageCount(), numStyle);
                cellNumber(row, 4, s.getPackageCount(), numStyle);  // 총입고량 = 초기재고
                cellNumber(row, 5, s.getOutboundQty(), numStyle);
                cellNumber(row, 6, s.getRemainingQty(), numStyle);
            }

            // 컬럼 자동 너비
            for (int i = 0; i < titles.length; i++) {
                sheet.autoSizeColumn(i);
                // autoSizeColumn 이 한글 폭을 잘 못 잡아서 약간 추가 여유
                int w = sheet.getColumnWidth(i);
                sheet.setColumnWidth(i, Math.min(w + 1500, 12000));
            }

            // ===== 응답 헤더 (한글 파일명 처리) =====
            String filename = "현재재고_" + LocalDate.now() + ".xlsx";
            String encoded = URLEncoder.encode(filename, StandardCharsets.UTF_8).replace("+", "%20");
            response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
            response.setHeader("Content-Disposition", "attachment; filename*=UTF-8''" + encoded);

            workbook.write(response.getOutputStream());
            response.getOutputStream().flush();
        }
    }

    /* ===== 내부 헬퍼 ===== */

    private List<Inbound> loadStocks(Boolean onlyAvailable) {
        boolean filterAvailable = Boolean.TRUE.equals(onlyAvailable);
        List<Inbound> list = inboundService.findAll(null, null);
        if (filterAvailable) {
            list = list.stream()
                    .filter(i -> i.getRemainingQty() != null && i.getRemainingQty() > 0)
                    .collect(Collectors.toList());
        }
        // 1차: 날짜(YYYYMMDD) 오름차순  2차: NNN 오름차순
        list.sort(Comparator
                .comparing(StockController::extractDate)
                .thenComparingInt(StockController::extractSuffix));
        return list;
    }

    /** P-YYYYMMDD-NNN 형식에서 YYYYMMDD 부분 추출. 형식 다르면 "99999999" (뒤로) */
    private static String extractDate(Inbound i) {
        String id = i == null ? null : i.getInboundId();
        if (id == null) return "99999999";
        // "P-" 다음 8자
        if (id.length() < 10 || id.charAt(1) != '-') return "99999999";
        String mid = id.substring(2, Math.min(10, id.length()));
        // 숫자인지 간단 검증
        for (int c = 0; c < mid.length(); c++) {
            if (!Character.isDigit(mid.charAt(c))) return "99999999";
        }
        return mid;
    }

    /** P-YYYYMMDD-NNN 형식에서 마지막 NNN 을 정수로 추출. 없으면 MAX_VALUE (뒤로) */
    private static int extractSuffix(Inbound i) {
        String id = i == null ? null : i.getInboundId();
        if (id == null) return Integer.MAX_VALUE;
        int dash = id.lastIndexOf('-');
        if (dash < 0 || dash == id.length() - 1) return Integer.MAX_VALUE;
        try {
            return Integer.parseInt(id.substring(dash + 1));
        } catch (NumberFormatException e) {
            return Integer.MAX_VALUE;
        }
    }

    private static void setBorder(CellStyle style) {
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
    }

    private static void cell(Row row, int idx, Object value, CellStyle style) {
        Cell c = row.createCell(idx);
        c.setCellStyle(style);
        if (value == null) return;
        if (value instanceof Number n) c.setCellValue(n.doubleValue());
        else c.setCellValue(value.toString());
    }

    private static void cellNumber(Row row, int idx, Number value, CellStyle style) {
        Cell c = row.createCell(idx);
        c.setCellStyle(style);
        if (value != null) c.setCellValue(value.doubleValue());
    }
}
