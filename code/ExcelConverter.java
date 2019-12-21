import java.io.*;
import java.util.*;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;

/**
 * Excel�u�b�N�R���o�[�^�[
 * 2003�ȑO<->2007�ȍ~ ���ݕϊ�
 */
class ExcelConverter {
	/**
	 * �G���g���[�|�C���g
	 *@param args[0] ���샂�[�h u = 2003->2007 d = 2007->2003
	 *@param ���̓��[�N�u�b�N�t�@�C����
	 *@param �o�̓��[�N�u�b�N�t�@�C����
	 */
	public static void main(String args[]) {
		// �p�����[�^�[�`�F�b�N
		if (args.length != 3) {
			System.out.println("�p�����[�^�[�G���[�ł��B");
			return;
		}
		String mode = args[0];
		if (!mode.equals("u") && !mode.equals("d")) {
			System.out.println("���샂�[�h�� d �܂��� u �Ŏw�肵�܂��B");
			return;
		}
		// �����J�n
		// Excel�u�b�N���I�[�v��
		Workbook workBook = null;
		try {
			if (mode.equals("u") {
				workBook = new HSSFWorkbook(new FileInputStream(args[1]));
			}
			else {
				workBook = new XSSFWorkbook(new FileInputStream(args[1]));
			}
		}
		catch (Exception e) {
			System.out.println("���̓u�b�N���J���܂���B" + e.ToStrig());
		}
		
	}
}

