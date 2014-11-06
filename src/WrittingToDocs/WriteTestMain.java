

package WrittingToDocs;

import java.io.IOException;
import java.util.logging.Level;
import java.util.logging.Logger;



public class WriteTestMain {
    
    public static int rand(){  
        return (int) (Math.random()*100);
    }
    
    public static void main(String[] args){
        int ha = 500;
        WriteToWord w = new WriteToWord(ha);
        
        int[] key = new int[ha];
        
        for(int x = 0; x < ha; x++){
            key[x] = x+1;
        }
        
        for(int c = 0; c < ha; c++){
            w.addTo(c, 0, key[c]);
            
            w.addTo(c, 1, rand());
            
            w.addTo(c, 2, rand());
            
            w.addTo(c, 3, rand());
            
            w.addTo(c, 4, rand());
                       
            w.addTo(c, 5, rand());
            
            w.addTo(c, 6, rand());
            
            w.addTo(c, 7, rand());
            
            w.addTo(c, 8, rand()); 
        }
        
       
        /*
        for(int c = 0; c < 100; c++){
            w.addTo(c, 0, key[c]);
            
            w.addTo(c, 1, (key[c] + c + 010));
            
            w.addTo(c, 2, (key[c] + c));
            
            w.addTo(c, 3, (key[c] + c + 010));
            
            w.addTo(c, 4, (key[c] + c));
                       
            w.addTo(c, 5, (key[c] + c + 010));
            
            w.addTo(c, 6, (key[c] + c));
            
            w.addTo(c, 7, (key[c] + c + 010));
            
            w.addTo(c, 8, (key[c] + c)); 
        }
        */
        try {
            w.WriteToMs();
        } catch (IOException ex) {
            Logger.getLogger(WriteTestMain.class.getName()).log(Level.SEVERE, null, ex);
        }
        
        int [][] test = new int [2][2];
        test[0][0]=1;
        test[0][1]=2;
        test[1][0]=3;
        test[1][1]=4;
        System.out.println(test[0][0] + " " + test[0][1]);
        System.out.println(test[1][0] + " " + test[1][1]);
        
        
        
    }
    
}
