import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;

import javax.swing.JButton;
import javax.swing.JCheckBox;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.JTextField;







import javax.swing.WindowConstants;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
public class scouting extends JFrame implements ActionListener{
	public static void main(String[] args){
		scouting scout = new scouting();
	}
	JLabel Lteam, Lmatch, Lauton,Lmoved,LHigh,LLow,LMultiBall,LAssists,LTeleop,LHits,LMisses,LCatches,LFail,LHighTele,LLowTele,LTrussTele,LHighTeleM,LLowTeleM,LTrussTeleM;
	JButton Submit = new JButton("Submit");
	JButton Clear = new JButton("Clear");
	JCheckBox Bmoved,BHighAuto, BLowAuto,BFail,BMultipleBalls;
	JTextField TTeam,TMatch,TTrussHit,TTrussmiss,TCatch,TAssists,THighHit,THighMiss,TLowHit,TLowMiss;
	public scouting(){
		JFrame window = new JFrame("Scouting");
		window.setDefaultCloseOperation(WindowConstants.EXIT_ON_CLOSE);
		JPanel components = new JPanel();
		window.setVisible(true);
		Lteam = new JLabel("Team Number:");
		Lmatch = new JLabel("  Match Number:");
		Lauton = new JLabel("   Auton:");
		Lmoved = new JLabel("  Moved");
		LHigh = new JLabel("  High Goal");
		LLow = new JLabel("  Low Goal");
		LMultiBall = new JLabel("  Multiple Balls");
		LTeleop = new JLabel("    Teleop:");
		LAssists = new JLabel("   Assists");
		LHits = new JLabel("   Hits");
		LMisses = new JLabel("   Misses");
		LCatches = new JLabel("   Catches");
		LFail = new JLabel("  Fail");
		LHighTele = new JLabel("   High Goal");
		LLowTele = new JLabel("   Low Goal");
		LTrussTele = new JLabel("   Truss Shots");
		LHighTeleM = new JLabel("   High Goal");
		LLowTeleM = new JLabel("   Low Goal");
		LTrussTeleM = new JLabel("   Truss Shots");
		Bmoved = new JCheckBox();
		BHighAuto = new JCheckBox();
		BLowAuto = new JCheckBox();
		BFail = new JCheckBox();
		TTeam = new JTextField("", 3);
		TMatch = new JTextField("", 3);
		TTrussHit = new JTextField("", 3);
		TTrussmiss = new JTextField("", 3);
		TCatch = new JTextField("", 3);
		TAssists = new JTextField("", 3);
		THighHit = new JTextField("", 3);
		THighMiss = new JTextField("", 3);
		TLowHit = new JTextField("", 3);
		TLowMiss = new JTextField("", 3);
		BMultipleBalls = new JCheckBox();
		Submit.addActionListener(this);
		Clear.addActionListener(this);
		components.add(Lteam);
		components.add(TTeam);
		components.add(Lmatch);
		components.add(TMatch);
		components.add(Lauton);
		components.add(Lmoved);
		components.add(Bmoved);
		components.add(LHigh);
		components.add(BHighAuto);
		components.add(LLow);
		components.add(BLowAuto);
		components.add(LMultiBall);
		components.add(BMultipleBalls);
		components.add(LTeleop);
		components.add(LAssists);
		components.add(TAssists);
		components.add(LHits);
		components.add(LHighTele);
		components.add(THighHit);
		components.add(LLowTele);
		components.add(TLowHit);
		components.add(LTrussTele);
		components.add(TTrussHit);
		components.add(LMisses);
		components.add(LHighTeleM);
		components.add(THighMiss);
		components.add(LLowTeleM);
		components.add(TLowMiss);
		components.add(LTrussTeleM);
		components.add(TTrussmiss);
		components.add(LCatches);
		components.add(TCatch);
		components.add(LFail);
		components.add(BFail);
		components.add(Submit);
		components.add(Clear);
		window.add(components);
		window.pack();


	}
	public void actionPerformed(ActionEvent e) {
		// TODO Auto-generated method stub
		if(e.getActionCommand().equals("Submit")){
			String team = TTeam.getText();
			String match = TMatch.getText();
			String assists = TAssists.getText();
			String hightelehit = THighHit.getText();
			String lowtelehit = TLowHit.getText();
			String hightelemiss = THighMiss.getText();
			String lowtelemiss = TLowMiss.getText();
			String trusshit = TTrussHit.getText();
			String trussmiss = TTrussmiss.getText();
			String catches = TCatch.getText();
			Boolean autonmove = Bmoved.isSelected();
			Boolean autonhigh = BHighAuto.isSelected();
			Boolean autonlow = BLowAuto.isSelected();
			Boolean multipleautoshots = BMultipleBalls.isSelected();
			Boolean failure = BFail.isSelected();
			int autonmoveint = autonmove? 1:0;
			int autonlowint = autonlow? 1:0;
			int autonhighint = autonhigh? 1:0;
			int multipleautoshotsint = multipleautoshots? 1:0;
			int failureint = failure? 1:0;
			try {
				FileInputStream input = new FileInputStream(new File("workbook.xls"));
				HSSFWorkbook wb = new HSSFWorkbook(input);
				HSSFSheet s = wb.getSheetAt(0);
				int lastrow = s.getLastRowNum();
				String rownumber = Integer.toString(lastrow+1);
				Map<String, Object[]> data = new TreeMap<String, Object[]>();
				data.put(rownumber, new Object[]{team,match,autonmoveint,autonhighint,autonlowint,multipleautoshotsint,hightelehit,lowtelehit,assists,hightelemiss,lowtelemiss,trusshit,trussmiss,catches,failureint});
				Set<String> keyset = data.keySet();
				int rownum = Integer.parseInt(rownumber);
				for (String key : keyset)
				{
					Row row = s.createRow(rownum);
					Object [] objArr = data.get(key);
					int cellnum = 0;
					for (Object obj : objArr)
					{
						Cell cell = row.createCell(cellnum++);
						if(obj instanceof String)
							cell.setCellValue((String)obj);
						else if(obj instanceof Integer)
							cell.setCellValue((Integer)obj);
					}
				}
				FileOutputStream output = new FileOutputStream("workbook.xls");
				wb.write(output);
			} catch (FileNotFoundException e1) {
				// TODO Auto-generated catch block

				try {
					FileOutputStream out = new FileOutputStream("workbook.xls");
					HSSFWorkbook wb = new HSSFWorkbook();
					HSSFSheet s = wb.createSheet("ScoutingData");
					Map<String, Object[]> data = new TreeMap<String, Object[]>();
					data.put("0", new Object[] {"Team", "Match", "Moves in Autonomous","High Auton Hits", "Low Auton Hits","Multiball Auton","Teleop High hits","Teleop low hits",
							"Assists","Teleop High misses","Teleop Low misses","Successful trusses","Unsuccessful trusses","Truss catches","Failure of bot"});
					int rownum = 0;
					Set<String> keyset = data.keySet();
					for (String key : keyset)
					{
						Row row = s.createRow(rownum++);
						Object [] objArr = data.get(key);
						int cellnum = 0;
						for (Object obj : objArr)
						{
							Cell cell = row.createCell(cellnum++);
							if(obj instanceof String)
								cell.setCellValue((String)obj);
							else if(obj instanceof Integer)
								cell.setCellValue((Integer)obj);
						}
					}
					JOptionPane.showMessageDialog(null, "Workbook not found, blank one created.  Press submit again to write data.");
					wb.write(out);
				} catch (FileNotFoundException e2) {
					// TODO Auto-generated catch block
					e2.printStackTrace();
				} catch (IOException e2) {
					// TODO Auto-generated catch block
					e2.printStackTrace();
				}

			} catch (IOException e1) {
				// TODO Auto-generated catch block
				e1.printStackTrace();
			}






		}
		else if(e.getActionCommand().equals("Clear")){
			TTeam.setText("");
			TTrussHit.setText("");
			TTrussmiss.setText("");
			TCatch.setText("");
			TAssists.setText("");
			THighHit.setText("");
			THighMiss.setText("");
			TLowHit.setText("");
			TLowMiss.setText("");
			Bmoved.setSelected(false);
			BHighAuto.setSelected(false);
			BLowAuto.setSelected(false);
			BFail.setSelected(false);
			BMultipleBalls.setSelected(false);
			JOptionPane.showMessageDialog(null, "All unsaved entries have been cleared.");
			
			
		}
	}


}