import java.io.File;

/**
 * Created by a on 2017-02-03.
 */
public class AppMain {
    public static void main(String[] args){
        System.out.println("[ :: Richware Systems EXCELtoWORD Converter :: ]");
        if(args.length < 2){
            System.out.println("입력 형식이 올바르지 않습니다.");
            System.out.println("도움말을 보려면 more help를 옵션으로 실행하십시오.");
            return;
        }else{
            if(args[0].equals("more")){
                switch (args[1]){
                    case "help":
                        System.out.println("입력파일.xls 출력파일.doc 버전명 프로젝트제목");
                        System.out.println("위와 같은 형식으로 입력하여 변환을 수행합니다.");
                        System.out.println("엑셀은 xls 형식이며, 워드는 doc입니다.");
                        return;
                    case "author":
                        System.out.println("Richware Systems : Ham");
                        return;
                    default:
                        System.out.println(args[1] + " : 알 수 없는 옵션입니다.");
                        return;
                }
            }else{
                if(args[0].contains(".xlsx")){
                    System.out.println(".xlsx 파일은 지원하지 않습니다. 저장 시 옵션을 변경하여 저장하세요.");
                    return;
                }
                if(args.length < 4){
                    System.out.println("프로젝트 제목을 입력하세요.");
                    System.out.println("도움말을 보려면 more help를 옵션으로 실행하십시오.");
                    return;
                }
                String title = "";
                for(int i = 3; i < args.length; i++) title += args[i] + " ";

                XWConverter xwConverter = new XWConverter(args[0], args[1], args[2],title.trim());
                xwConverter.execute();
                return;
            }
        }

    }
}
