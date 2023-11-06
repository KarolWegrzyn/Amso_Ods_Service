import org.odftoolkit.odfdom.doc.OdfSpreadsheetDocument;
import org.odftoolkit.odfdom.doc.table.OdfTable;
//import org.odftoolkit.odfdom.doc.OdfTextDocument;
//import org.odftoolkit.odfdom.changes.Table;
//import org.odftoolkit.odfdom.doc.OdfDocument.*;
//import java.util.List;

public class App {

    public static void main(String[] args) {
        try 
        {
        	OdfSpreadsheetDocument spreadsheetDoc = OdfSpreadsheetDocument.loadDocument("plik.ods");
            OdfTable tab = spreadsheetDoc.getTableList().get(0);

            int liczbaWierszy = tab.getRowCount();
            //System.out.println(tab.getCellByPosition(1,1).getStringValue());
            System.out.println("Liczba wierszy:" + liczbaWierszy);
            System.out.println("Liczba kolumn:" + tab.getColumnCount());
            System.out.println();
            //System.out.println();
            
            for(int i=0; i<tab.getRowCount() ;i++) //w osi y 
            {
            	for(int j=0; j<tab.getColumnCount();j++) //w osi x
            	{	
            		if(j==2 && i>0) 
            		{
            			
            			long wartZmien = Long.parseLong(tab.getCellByPosition(j, i).getStringValue());
            			
            			if(wartZmien !=0)
            			{
            				String id_Iai = tab.getCellByPosition(1, i).getStringValue();
            				
            				tab.getCellByPosition(3,i).setStringValue(id_Iai); //modyfikacja wartosci konkretnej kom
            			}
            			else if(wartZmien==0)
            			{
            				long wZ = Long.parseLong(tab.getCellByPosition(j, i+1).getStringValue());
            				
            				if(wZ!=0)
            				{
                				String nr_Zapas = tab.getCellByPosition(0, i+1).getStringValue();
                				tab.getCellByPosition(3,i).setStringValue(nr_Zapas);
            				}
            				else
            				{
            					boolean czyZero = true;
            					String nr_Zapas;
            					int k = 0;
            					
            					while(czyZero)
            					{
            						long wartZmiennej = Long.parseLong(tab.getCellByPosition(j, i+k).getStringValue());
            						
            						if(wartZmiennej != 0)
            						{
            							czyZero = false;
            							nr_Zapas = tab.getCellByPosition(0, i+k).getStringValue();
            							tab.getCellByPosition(3,i).setStringValue(nr_Zapas);
            						}
            						else if(wartZmiennej == 0)
            						{
            							k++;
            						}     							
            					}
            				}
            			}
            		}
            		System.out.print(tab.getCellByPosition(j, i).getStringValue() + "\t");
            	}
        		if(i==0)
        			System.out.println();
        		
            	System.out.println();
            }
            spreadsheetDoc.save("wynik_programu.ods");
            spreadsheetDoc.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
        
    }
}