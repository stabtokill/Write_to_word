package WrittingToDocs;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

/**
 *
 * @author Brian_Silewski
 */
public class WriteToWord {

    private int[][] results;
    private int bs; //default 104
    private int ss; //default 100

    //setting size of arrays and table
    public WriteToWord(int a) {
        ss = a;
        bs = a + 4;
        results = new int[a][9];
    }

    //basicaly is writes everything into a word document
    public void WriteToMs() throws IOException {
        XWPFDocument document = new XWPFDocument();

        XWPFTableRow[] tableRow = new XWPFTableRow[bs + 1];

        XWPFTable tableOne = document.createTable();
        XWPFTableRow tableOneRowOne = tableOne.getRow(0);

        tableOneRowOne.getCell(0).setText("Header1");

        //sets top headers
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

        //sets headers
        for (int xx = 1; xx < 9; xx++) {
            if (xx % 2 == 0) {
                tableOneRowThree.getCell(xx).setText("Time");
            } else {
                tableOneRowThree.getCell(xx).setText("Index");
            }
        }

        //creates many objects for more rows
        for (int a = 4; a <= bs; a++) {
            tableRow[a] = tableOne.createRow();
        }

        //set all the values from array to table
        for (int q = 4; q < bs; q++) {
            for (int u = 0; u < 9; u++) {
                tableRow[q].getCell(u).setText("" + results[(q - 4)][u]);
            }

        }

        //adds the values for the average time, always last row in table
        for (int xx = 1; xx < 9; xx++) {
            if (xx % 2 == 0) {
                tableRow[bs].getCell(xx).setText("" + aveg(xx));
            } else {
                tableRow[bs].getCell(xx).setText("Avg Time");
            }
        }
        
        //creates the file directory for the file to be saved at
        String namePath;
        namePath = setup();
        namePath += "/output doc";

        File f = new File(namePath);

        //test to see if it is already there, if not, it creates it
        if (!f.exists()) {
            boolean success;
            success = (new File(namePath)).mkdirs();
            if (!success) {
                System.out.println("\nFailed to create new directory, quiting");
                System.out.close();
            }
        } else {
            System.out.println("\nFile 1 is already created");
        }

        //saves the file
        FileOutputStream outStream = new FileOutputStream(namePath + "\\test.doc");
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

    //finds the desktop
    public String setup() {
        String desktopPath = System.getProperty("user.home") + "/Desktop";
        System.out.print(desktopPath.replace("\\", "/"));
        return desktopPath;
    }

}
