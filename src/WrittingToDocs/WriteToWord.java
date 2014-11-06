package WrittingToDocs;

import java.io.FileOutputStream;
import java.io.IOException;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

/**
 *
 * @author brian_000
 */
public class WriteToWord {

    private int[][] results;
    private int bs; //default 104
    private int ss; //default 100
    
    //setting size of arrays and table
    public WriteToWord(int a){
        ss = a;
        bs = a+4;
        results = new int[a][9];
    }

    //basicaly is writes everything into a word document
    public void WriteToMs() throws IOException {
        XWPFDocument document = new XWPFDocument();

        XWPFTableRow[] tableRow = new XWPFTableRow[bs+1];

        XWPFTable tableOne = document.createTable();
        XWPFTableRow tableOneRowOne = tableOne.getRow(0);

        tableOneRowOne.getCell(0).setText("Header1");

        for (int xx = 1; xx < 9; xx++) {
            tableOneRowOne.addNewTableCell().setText("                 " + xx);
        }

        tableOneRowOne.getCell(3).removeParagraph(0);
        tableOneRowOne.getCell(7).removeParagraph(0);
        tableOneRowOne.getCell(3).setText("                SS");
        tableOneRowOne.getCell(7).setText("                BS");

        XWPFTableRow tableOneRowTwo = tableOne.createRow();

        tableOneRowTwo.getCell(2).setText("Unsorted");
        tableOneRowTwo.getCell(4).setText("Sorted");
        tableOneRowTwo.getCell(6).setText("Unsorted");
        tableOneRowTwo.getCell(8).setText("Sorted");

        XWPFTableRow tableOneRowThree = tableOne.createRow();

        tableOneRowThree.getCell(0).setText("Key");

        for (int xx = 1; xx < 9; xx++) {
            if (xx % 2 == 0) {
                tableOneRowThree.getCell(xx).setText("Time");
            } else {
                tableOneRowThree.getCell(xx).setText("Index");
            }
        }

        for (int a = 4; a <= bs; a++) {
            tableRow[a] = tableOne.createRow();
        }

        for (int q = 4; q < bs; q++) {
            for (int u = 0; u < 9; u++) {
                tableRow[q].getCell(u).setText("" + results[(q - 4)][u]);
            }

        }

        for (int xx = 1; xx < 9; xx++) {
            if (xx % 2 == 0) {
                tableRow[bs].getCell(xx).setText("" + aveg(xx));
            } else {
                tableRow[bs].getCell(xx).setText("Avg Time");
            }
        }

        FileOutputStream outStream = new FileOutputStream("C:\\Users\\brian_000\\Desktop\\create tests\\test.doc");
        document.write(outStream);
        outStream.close();
    }

    //returns the average
    public double aveg(int cell) {
        double y = 0;
        for (int q = 0; q < ss; q++) {
            y += results[q][cell];
        }
        y /= ss;
        return y;
    }

    //adds all the elements to the array
    public void addTo(int r, int c, int ans) {
        results[r][c] = ans;
    }

}
