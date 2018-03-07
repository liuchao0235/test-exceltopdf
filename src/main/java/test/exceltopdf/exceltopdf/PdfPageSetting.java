package test.exceltopdf.exceltopdf;

import com.itextpdf.text.PageSize;
import com.itextpdf.text.Rectangle;


public class PdfPageSetting {
	private float marginLeft = 10f;
	private float marginRight = 10f;
	private float marginTop = 10f;
	private float marginBottom = 10f;
	private Rectangle pageSize = PageSize.A4;
	public float getMarginLeft() {
		return marginLeft;
	}
	public void setMarginLeft(float marginLeft) {
		this.marginLeft = marginLeft;
	}
	public float getMarginRight() {
		return marginRight;
	}
	public void setMarginRight(float marginRight) {
		this.marginRight = marginRight;
	}
	public float getMarginTop() {
		return marginTop;
	}
	public void setMarginTop(float marginTop) {
		this.marginTop = marginTop;
	}
	public float getMarginBottom() {
		return marginBottom;
	}
	public void setMarginBottom(float marginBottom) {
		this.marginBottom = marginBottom;
	}
	public Rectangle getPageSize() {
		return pageSize;
	}
	public void setPageSize(Rectangle pageSize) {
		this.pageSize = pageSize;
	}
	
}
