import java.io.*;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.List;

import jxl.Sheet;
import jxl.Workbook;
import jxl.WorkbookSettings;
import jxl.read.biff.BiffException;
import word.api.interfaces.IDocument;
import word.w2004.Document2004;
import word.w2004.elements.*;
import word.w2004.elements.tableElements.TableEle;
import word.w2004.style.Font;

/**
 * Created by a on 2017-02-03.
 */
public class XWConverter {

    private boolean succFlag = true;
    private List<ExcelTable> tableList = null;
    private IDocument myDoc = null;
    private String origin = null;
    private String dest = null;
    private String subText = null;
    private String versionName = null;
    private Workbook workbook = null;
    private int apiCount = 0;

    public XWConverter(String originPath, String destPath, String v, String sub){
        tableList = new ArrayList<ExcelTable>();
        myDoc = new Document2004();
        subText = sub;
        versionName = v;
        origin = originPath;
        dest = destPath;

        WorkbookSettings setting = new WorkbookSettings();
        setting.setEncoding("EUC-KR");
        File file = new File(originPath);

        try {
            // 엑셀파일 워크북 객체 생성
            workbook = Workbook.getWorkbook(file, setting);
        } catch (BiffException e) {
            System.out.println("손상된 문서이거나 지원하지 않는 버전입니다.");
        } catch (IOException e) {
            System.out.println("파일이 존재하지 않습니다.");
        }

    }

    private void printTitle(){
        Calendar c = Calendar.getInstance();
        String date = c.get(Calendar.YEAR) + "-" + (c.get(Calendar.MONTH) + 1) + "-" + c.get(Calendar.DAY_OF_MONTH);
        myDoc.addEle(BreakLine.times(8).create());
        myDoc.addEle(Paragraph.withPieces(ParagraphPiece.with("프로토콜 연동 문서").withStyle().font(Font.CENTURY_GOTHIC).fontSize("40").bold().create()).create());
        myDoc.addEle(Paragraph.withPieces(ParagraphPiece.with(subText).withStyle().font(Font.CENTURY_GOTHIC).fontSize("16").bold().create()).create());
        myDoc.addEle(BreakLine.times(26).create());
        myDoc.addEle(Paragraph.withPieces(ParagraphPiece.with("최초 작성일 : " + date).withStyle().font(Font.CENTURY_GOTHIC).fontSize("12").bold().create()).create());
        myDoc.addEle(Paragraph.withPieces(ParagraphPiece.with("버전 " + versionName).withStyle().font(Font.CENTURY_GOTHIC).fontSize("12").bold().create()).create());
        myDoc.addEle(PageBreak.create());
    }

    private void parseAndWrite(){
        Sheet sheet = workbook.getSheet(0);
        int endIdx = sheet.getColumn(1).length;

        System.out.println("테이블 파싱을 시작합니다.");

        for(int i = 0; i < endIdx; i++) {
            if(sheet.getCell(0, i).getContents().equals("API 명")) { // Recognizing every single API table as a unit
                apiCount++;
                String apiName = sheet.getCell(1, i).getContents();
                String expText = sheet.getCell(1, i + 1).getContents();
                ArrayList<ExcelRow> params = new ArrayList<ExcelRow>();
                ArrayList<ExcelRow> returns = new ArrayList<ExcelRow>();

                int j;

                for(j = i + 3; !sheet.getCell(0, j).getContents().equals("CALL SAMPLE"); j++) {
                    params.add(new ExcelRow(new String[]{
                            sheet.getCell(1, j).getContents(),
                            sheet.getCell(2, j).getContents(),
                            sheet.getCell(3, j).getContents(),
                            sheet.getCell(4, j).getContents(),
                            sheet.getCell(5, j).getContents()}));
                }

                String sample = sheet.getCell(1, j).getContents();
                String ret = sheet.getCell(1, j + 1).getContents();

                for(j += 3; !sheet.getCell(1, j).getContents().trim().equals(""); j++) {
                    returns.add(new ExcelRow(new String[]{
                            sheet.getCell(1, j).getContents(),
                            sheet.getCell(2, j).getContents(),
                            sheet.getCell(3, j).getContents(),
                            sheet.getCell(4, j).getContents(),
                            sheet.getCell(5, j).getContents()}));
                    if(j == endIdx - 1) break;
                }

                ExcelTable excelTable = new ExcelTable(apiName, expText, sample, ret, params, returns);
                tableList.add(excelTable);
            }
        }

        System.out.println("총 " + apiCount + "개의 API 테이블 인식이 완료되었습니다.");

        for(int i = 0; i < tableList.size(); i++){ // API 테이블 작성 루틴
            myDoc.addEle(Paragraph.withPieces(ParagraphPiece.with((i + 1) + ". " + tableList.get(i).api).withStyle().font(Font.CENTURY_GOTHIC).fontSize("10").bold().create()).create());
            myDoc.addEle(Paragraph.withPieces(ParagraphPiece.with(" A. 접속 " + tableList.get(i).sam).withStyle().font(Font.CENTURY_GOTHIC).fontSize("10").create()).create());
            myDoc.addEle(Paragraph.withPieces(ParagraphPiece.with(" B. 필요 파라미터").withStyle().font(Font.CENTURY_GOTHIC).fontSize("10").create()).create());

            Table param = new Table();
            param.addTableEle(TableEle.TH,
                    Paragraph.withPieces(ParagraphPiece.with("파라미터").withStyle().fontSize("10").create()).create(),
                    Paragraph.withPieces(ParagraphPiece.with("설명").withStyle().fontSize("10").create()).create(),
                    Paragraph.withPieces(ParagraphPiece.with("필수 여부").withStyle().fontSize("10").create()).create(),
                    Paragraph.withPieces(ParagraphPiece.with("비고").withStyle().fontSize("10").create()).create()
            );

            for(int e = 0; e < tableList.get(i).listParams.size(); e++) {
                param.addTableEle(TableEle.TD,
                        Paragraph.withPieces(ParagraphPiece.with(tableList.get(i).listParams.get(e).list[0]).withStyle().fontSize("10").create()).create(),
                        Paragraph.withPieces(ParagraphPiece.with(tableList.get(i).listParams.get(e).list[4]).withStyle().fontSize("10").create()).create(),
                        Paragraph.withPieces(ParagraphPiece.with(tableList.get(i).listParams.get(e).list[2]).withStyle().fontSize("10").create()).create(),
                        Paragraph.withPieces(ParagraphPiece.with(tableList.get(i).listParams.get(e).list[3]).withStyle().fontSize("10").create()).create()
                );
            }

            myDoc.addEle(param);

            myDoc.addEle(BreakLine.times(1).create());

            myDoc.addEle(Paragraph.withPieces(ParagraphPiece.with(" C. 리턴 결과").withStyle().font(Font.CENTURY_GOTHIC).fontSize("10").create()).create());

            Table returnTable = new Table();
            returnTable.addTableEle(TableEle.TH,
                    Paragraph.withPieces(ParagraphPiece.with("1 depth").withStyle().fontSize("10").create()).create(),
                    Paragraph.withPieces(ParagraphPiece.with("2 depth").withStyle().fontSize("10").create()).create(),
                    Paragraph.withPieces(ParagraphPiece.with("3 depth").withStyle().fontSize("10").create()).create(),
                    Paragraph.withPieces(ParagraphPiece.with("설명").withStyle().fontSize("10").create()).create(),
                    Paragraph.withPieces(ParagraphPiece.with("비고").withStyle().fontSize("10").create()).create()
            );

            for(int e = 0; e < tableList.get(i).listReturns.size(); e++) {
                returnTable.addTableEle(TableEle.TD,
                        Paragraph.withPieces(ParagraphPiece.with(tableList.get(i).listReturns.get(e).list[0]).withStyle().fontSize("10").create()).create(),
                        Paragraph.withPieces(ParagraphPiece.with(tableList.get(i).listReturns.get(e).list[3]).withStyle().fontSize("10").create()).create(),
                        Paragraph.withPieces(ParagraphPiece.with(tableList.get(i).listReturns.get(e).list[3]).withStyle().fontSize("10").create()).create(),
                        Paragraph.withPieces(ParagraphPiece.with(tableList.get(i).listReturns.get(e).list[4]).withStyle().fontSize("10").create()).create(),
                        Paragraph.withPieces(ParagraphPiece.with(tableList.get(i).listReturns.get(e).list[3]).withStyle().fontSize("10").create()).create()
                );
            }

            myDoc.addEle(returnTable);

            myDoc.addEle(BreakLine.times(1).create());
        }

    }

    private void writeFile(){
        File fileObj = new File(dest);
        PrintWriter writer = null;
        try {
            writer = new PrintWriter(fileObj);
        } catch (FileNotFoundException e) {
            System.out.println("파일이 사용 중이거나 존재하지 않습니다.");
            succFlag = false;
            return;
        }
        String myWord = myDoc.getContent();

        writer.println(myWord);
        writer.close();
    }

    public void execute(){
        printTitle();
        parseAndWrite();
        writeFile();

        if(succFlag) System.out.println(dest + " - 변환 작업이 완료되었습니다.");
        else System.out.println(dest + " - 변환 작업에 실패하였습니다.");
    }

}