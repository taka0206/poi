import java.io.*;
import java.util.*;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.HSSFColor;

public class pict{
	public static void main(String[] args){
		System.out.println("PICTURE_TYPE_EMF  " +  HSSFWorkbook.PICTURE_TYPE_EMF);
		System.out.println("PICTURE_TYPE_WMF  " +  HSSFWorkbook.PICTURE_TYPE_WMF);
		System.out.println("PICTURE_TYPE_PICT " +  HSSFWorkbook.PICTURE_TYPE_PICT);
		System.out.println("PICTURE_TYPE_JPEG " +  HSSFWorkbook.PICTURE_TYPE_JPEG);
		System.out.println("PICTURE_TYPE_DIB  " +  HSSFWorkbook.PICTURE_TYPE_DIB);
		System.out.println("HSSFColor.AQUA.index " + HSSFColor.AQUA.index);

	}
}
