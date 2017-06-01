import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;

import org.apache.poi.POIXMLDocument;
import org.apache.poi.POIXMLTextExtractor;
import org.apache.poi.hwpf.extractor.WordExtractor;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.extractor.XWPFWordExtractor;

/**
 * POI 读取 word 2003 和 word 2007 中文字内容的测试类<br />
 * @createDate 2009-07-25
 * @author Carl He
 */
public class Test {
	public static void main(String[] args) {
		try {
			////word 2003： 图片不会被读取
			/*InputStream is = new FileInputStream(new File("files\\2003.doc"));
			WordExtractor ex = new WordExtractor(is);//is是WORD文件的InputStream 
			String text2003 = ex.getText();
			System.out.println(text2003);*/

			//word 2007 图片不会被读取， 表格中的数据会被放在字符串的最后
//			OPCPackage opcPackage = POIXMLDocument.openPackage("files\\2007.docx");
			OPCPackage opcPackage = POIXMLDocument.openPackage("files\\test.docx");
			
			POIXMLTextExtractor extractor = new XWPFWordExtractor(opcPackage);
			StringBuilder sb = new StringBuilder();
			sb.append(extractor.getText());
			int index = sb.indexOf("选择");
			System.out.println(index);
			
			//进入选择题区域
			while(index!=-1){
				int i = 1;
				//题干的起始位置
				int numStart = 0;
				int numEnd = 0;
				while(numStart!=-1){
					//获取题干
					numStart = sb.indexOf("("+i+")",index+1);
					numEnd = sb.indexOf("。",numStart+1);
					String ques = sb.substring(numStart, numEnd).trim();
					System.out.println(ques);
					//获取题支
					numStart = sb.indexOf("("+(++i)+")",numStart);//下一个题干的位置
					if(numStart==-1){
						int AIndex = sb.indexOf("A",numEnd);
						int end = sb.indexOf("应用题",numEnd);
						String options = sb.substring(numEnd+1, end-1).trim();
						System.out.println(options);
					}else{
						int AIndex = sb.indexOf("A",numEnd);
						String options = sb.substring(AIndex,numStart).trim();
						options = options.trim();
						System.out.println(options);
					}
				}
				index = sb.indexOf("选择",index+1);
				System.out.println(index);
			}
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
}