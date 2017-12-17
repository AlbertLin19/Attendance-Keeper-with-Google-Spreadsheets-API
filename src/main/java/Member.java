/**
 * This class just holds the necessary information that a team member needs for identification on a spreadsheet.
 * @author Albert Lin
 *
 */
public class Member {
	private String name;
	private int rowNumber;
	
	public Member(String nameIn, int rowNumberIn) {
		name = nameIn;
		rowNumber = rowNumberIn;
	}
	
	public String getName() {
		return name;
	}
	
	public int getRowNumber() {
		return rowNumber;
	}
	
}
