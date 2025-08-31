package org.wordinator.xml2docx.generator;

import org.apache.poi.xwpf.usermodel.XWPFTable.XWPFBorderType;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STBorder;
import org.apache.xmlbeans.XmlCursor;
import org.apache.xmlbeans.XmlObject;

/**
 * Holds a set of table border styles
 */
public class TableBorderStyles {

  // Default border type is set by the @borderstyle or @framestyle attribute.
  // By default there are no explicit borders.
  XWPFBorderType defaultBorderType = null;
  XWPFBorderType topBorder = null;
  XWPFBorderType bottomBorder = null;
  XWPFBorderType leftBorder = null;
  XWPFBorderType rightBorder = null;
  XWPFBorderType rowSepBorder = null;
  XWPFBorderType colSepBorder = null;

  String defaultColor = null;
  String topColor = null;
  String leftColor = null;
  String bottomColor = null;
  String rightColor = null;

  boolean hasBorderWidth = false;

  public TableBorderStyles(XWPFBorderType defaultBorderType,
                           XWPFBorderType topBorder,
                           XWPFBorderType bottomBorder,
                           XWPFBorderType leftBorder,
                           XWPFBorderType rightBorder) {
  }

  /**
   * Construct using specified border styles as the initial values.
   * @param parentBorderStyles Styles to be inherited from parent
   */
  public TableBorderStyles(TableBorderStyles parentBorderStyles) {
    defaultBorderType = parentBorderStyles.getDefaultBorderType();
    topBorder = parentBorderStyles.getTopBorder();
    bottomBorder = parentBorderStyles.getBottomBorder();
    leftBorder = parentBorderStyles.getLeftBorder();
    rightBorder = parentBorderStyles.getRightBorder();
    rowSepBorder = parentBorderStyles.getRowSepBorder();
    colSepBorder = parentBorderStyles.getColSepBorder();

    // Get default border colors from parent?
  }

  /**
   * Construct initial border styles from an element that may specify
   * border frame style attributes.
   * @param borderStyleSpecifier XML element that may specify frame style attributes (table, td)
   */
  public TableBorderStyles(XmlObject borderStyleSpecifier) {

    XmlCursor cursor = borderStyleSpecifier.newCursor();
    String tagname = cursor.getName().getLocalPart();
    String styleValue = null;
    String styleBottomValue= null;
    String styleTopValue= null;
    String styleLeftValue= null;
    String styleRightValue= null;

    String colorValue = null;
    String colorBottomValue= null;
    String colorTopValue= null;
    String colorLeftValue= null;
    String colorRightValue= null;

    // Issue 30: Also get the border color values.

    if ("table".equals(tagname)) {
      styleValue = cursor.getAttributeText(DocxConstants.QNAME_FRAMESTYLE_ATT);
      styleBottomValue= cursor.getAttributeText(DocxConstants.QNAME_FRAMESTYLE_BOTTOM_ATT);
      styleTopValue= cursor.getAttributeText(DocxConstants.QNAME_FRAMESTYLE_TOP_ATT);
      styleLeftValue= cursor.getAttributeText(DocxConstants.QNAME_FRAMESTYLE_LEFT_ATT);
      styleRightValue= cursor.getAttributeText(DocxConstants.QNAME_FRAMESTYLE_RIGHT_ATT);
    } else {
      styleValue = cursor.getAttributeText(DocxConstants.QNAME_BORDER_STYLE_ATT);
      styleBottomValue= cursor.getAttributeText(DocxConstants.QNAME_BORDER_STYLE_BOTTOM_ATT);
      styleTopValue= cursor.getAttributeText(DocxConstants.QNAME_BORDER_STYLE_TOP_ATT);
      styleLeftValue= cursor.getAttributeText(DocxConstants.QNAME_BORDER_STYLE_LEFT_ATT);
      styleRightValue= cursor.getAttributeText(DocxConstants.QNAME_BORDER_STYLE_RIGHT_ATT);

      colorValue = cursor.getAttributeText(DocxConstants.QNAME_BORDER_COLOR_ATT);
      colorBottomValue = cursor.getAttributeText(DocxConstants.QNAME_BORDER_COLOR_BOTTOM_ATT);
      colorTopValue = cursor.getAttributeText(DocxConstants.QNAME_BORDER_COLOR_TOP_ATT);
      colorLeftValue = cursor.getAttributeText(DocxConstants.QNAME_BORDER_COLOR_LEFT_ATT);
      colorRightValue = cursor.getAttributeText(DocxConstants.QNAME_BORDER_COLOR_RIGHT_ATT);
    }

    if (styleValue != null) {
      setDefaultBorderType(xwpfBorderType(styleValue));
    }

    if (styleBottomValue != null) {
      setBottomBorder(xwpfBorderType(styleBottomValue));
    }
    if (styleTopValue != null) {
      setTopBorder(xwpfBorderType(styleTopValue));
    }
    if (styleLeftValue != null) {
      setLeftBorder(xwpfBorderType(styleLeftValue));
    }
    if (styleRightValue != null) {
      setRightBorder(xwpfBorderType(styleRightValue));
    }

    if (colorValue != null) {
      setDefaultBorderColor(colorValue);
    }

    if (colorBottomValue != null) {
      setBottomColor(colorBottomValue);
    }
    if (colorTopValue != null) {
      setTopColor(colorTopValue);
    }
    if (colorLeftValue != null) {
      setLeftColor(colorLeftValue);
    }
    if (colorRightValue != null) {
      setRightColor(colorRightValue);
    }

    hasBorderWidth = cursor.getAttributeText(DocxConstants.QNAME_BORDER_WIDTH_BOTTOM_ATT) != null ||
      cursor.getAttributeText(DocxConstants.QNAME_BORDER_WIDTH_LEFT_ATT) != null ||
      cursor.getAttributeText(DocxConstants.QNAME_BORDER_WIDTH_RIGHT_ATT) != null ||
      cursor.getAttributeText(DocxConstants.QNAME_BORDER_WIDTH_TOP_ATT) != null;
  }

  public void setDefaultBorderColor(String colorValue) {
    this.defaultColor = colorValue;
    if (this.getBottomColor() == null) this.setBottomColor(colorValue);
    if (this.getTopColor() == null) this.setTopColor(colorValue);
    if (this.getLeftColor() == null) this.setLeftColor(colorValue);
    if (this.getRightColor() == null) this.setRightColor(colorValue);

  }

  public String getBottomColor() {
    return this.bottomColor;
  }

  public String getTopColor() {
    return this.topColor;
  }

  public String getLeftColor() {
    return this.leftColor;
  }

  public String getRightColor() {
    return this.rightColor;
  }

  public void setBottomColor(String colorValue) {
    this.bottomColor = colorValue;
  }

  public void setTopColor(String colorValue) {
    this.topColor = colorValue;
  }

  public void setLeftColor(String colorValue) {
    this.leftColor = colorValue;
  }

  public void setRightColor(String colorValue) {
    this.rightColor = colorValue;
  }

  public XWPFBorderType getDefaultBorderType() {
    return defaultBorderType;
  }

  public void setDefaultBorderType(XWPFBorderType defaultBorderType) {
    this.defaultBorderType = defaultBorderType;
    if (getBottomBorder() == null) setBottomBorder(defaultBorderType);
    if (getTopBorder() == null) setTopBorder(defaultBorderType);
    if (getLeftBorder() == null) setLeftBorder(defaultBorderType);
    if (getRightBorder() == null) setRightBorder(defaultBorderType);
  }

  public XWPFBorderType getTopBorder() {
    return topBorder;
  }

  public void setTopBorder(XWPFBorderType topBorder) {
    this.topBorder = topBorder;
  }

  public XWPFBorderType getBottomBorder() {
    return bottomBorder;
  }

  public STBorder.Enum getBottomBorderEnum() {
    return getBorderEnumForType(getBottomBorder());
  }
  public STBorder.Enum getTopBorderEnum() {
    return getBorderEnumForType(getTopBorder());
  }
  public STBorder.Enum getLeftBorderEnum() {
    return getBorderEnumForType(getLeftBorder());
  }
  public STBorder.Enum getRightBorderEnum() {
    return getBorderEnumForType(getRightBorder());
  }

  public STBorder.Enum getBorderEnumForType(XWPFBorderType type) {
    STBorder.Enum result = null;
    if (type != null) {
      result = stBorderType(type);
    }
    return result;
  }

  public void setBottomBorder(XWPFBorderType bottomBorder) {
    this.bottomBorder = bottomBorder;
  }

  public XWPFBorderType getLeftBorder() {
    return leftBorder;
  }

  public void setLeftBorder(XWPFBorderType leftBorder) {
    this.leftBorder = leftBorder;
  }

  public XWPFBorderType getRightBorder() {
    return rightBorder;
  }

  public void setRightBorder(XWPFBorderType rightBorder) {
    this.rightBorder = rightBorder;
  }

  public XWPFBorderType getRowSepBorder() {
    return rowSepBorder;
  }

  public void setRowSepBorder(XWPFBorderType rowSepBorder) {
    this.rowSepBorder = rowSepBorder;
  }

  public XWPFBorderType getColSepBorder() {
    return colSepBorder;
  }

  public void setColSepBorder(XWPFBorderType colSepBorder) {
    this.colSepBorder = colSepBorder;
  }

  /**
   * Determine if any borders are explicitly set
   * @return True if one or more borders have a defined style.
   */
  public boolean hasBorders() {
    boolean result =
      getDefaultBorderType() != null ||
      getBottomBorder() != null ||
      getTopBorder() != null ||
      getLeftBorder() != null ||
      getRightBorder() != null ||
      hasBorderWidth;
    return result;
  }

  /**
   * Get the XWPFBorderType for the specified STBorder value.
   * @param borderValue Border value (e.g., "wave").
   * @return Corresponding XWPFBorderType value or null if there is no corresponding value.
   */
  private XWPFBorderType xwpfBorderType(String borderValue) {
    STBorder.Enum borderStyle = STBorder.Enum.forString(borderValue);

    // There's not a direct correspondence between STBorder int values
    // and XWPFBorderType so just building a switch statement.
    XWPFBorderType xwpfType = null;
    switch (borderStyle.intValue()) {
    case STBorder.INT_DOT_DASH:
      xwpfType = XWPFBorderType.DOT_DASH;
      break;
    case STBorder.INT_DASH_SMALL_GAP:
      xwpfType = XWPFBorderType.DASH_SMALL_GAP;
      break;
    case STBorder.INT_DASH_DOT_STROKED:
      xwpfType = XWPFBorderType.DASH_DOT_STROKED;
      break;
    case STBorder.INT_DASHED:
      xwpfType = XWPFBorderType.DASHED;
      break;
    case STBorder.INT_DOT_DOT_DASH:
      xwpfType = XWPFBorderType.DOT_DOT_DASH;
      break;
    case STBorder.INT_DOTTED:
      xwpfType = XWPFBorderType.DOTTED;
      break;
    case STBorder.INT_DOUBLE:
      xwpfType = XWPFBorderType.DOUBLE;
      break;
    case STBorder.INT_DOUBLE_WAVE:
      xwpfType = XWPFBorderType.DOUBLE_WAVE;
      break;
    case STBorder.INT_INSET:
      xwpfType = XWPFBorderType.INSET;
      break;
    case STBorder.INT_NIL:
      xwpfType = XWPFBorderType.NIL;
      break;
    case STBorder.INT_NONE:
      xwpfType = XWPFBorderType.NONE;
      break;
    case STBorder.INT_OUTSET:
      xwpfType = XWPFBorderType.OUTSET;
      break;
    case STBorder.INT_SINGLE:
      xwpfType = XWPFBorderType.SINGLE;
      break;
    case STBorder.INT_THICK:
      xwpfType = XWPFBorderType.THICK;
      break;
    case STBorder.INT_THICK_THIN_LARGE_GAP:
      xwpfType = XWPFBorderType.THICK_THIN_LARGE_GAP;
      break;
    case STBorder.INT_THICK_THIN_MEDIUM_GAP:
      xwpfType = XWPFBorderType.THICK_THIN_MEDIUM_GAP;
      break;
    case STBorder.INT_THICK_THIN_SMALL_GAP:
      xwpfType = XWPFBorderType.THICK_THIN_SMALL_GAP;
      break;
    case STBorder.INT_THIN_THICK_LARGE_GAP:
      xwpfType = XWPFBorderType.THIN_THICK_LARGE_GAP;
      break;
    case STBorder.INT_THIN_THICK_MEDIUM_GAP:
      xwpfType = XWPFBorderType.THIN_THICK_MEDIUM_GAP;
      break;
    case STBorder.INT_THIN_THICK_SMALL_GAP:
      xwpfType = XWPFBorderType.THIN_THICK_SMALL_GAP;
      break;
    case STBorder.INT_THIN_THICK_THIN_LARGE_GAP:
      xwpfType = XWPFBorderType.THIN_THICK_THIN_LARGE_GAP;
      break;
    case STBorder.INT_THIN_THICK_THIN_MEDIUM_GAP:
      xwpfType = XWPFBorderType.THIN_THICK_THIN_MEDIUM_GAP;
      break;
    case STBorder.INT_THIN_THICK_THIN_SMALL_GAP:
      xwpfType = XWPFBorderType.THIN_THICK_THIN_SMALL_GAP;
      break;
    case STBorder.INT_THREE_D_EMBOSS:
      xwpfType = XWPFBorderType.THREE_D_EMBOSS;
      break;
    case STBorder.INT_THREE_D_ENGRAVE:
      xwpfType = XWPFBorderType.THREE_D_ENGRAVE;
      break;
    case STBorder.INT_TRIPLE:
      xwpfType = XWPFBorderType.TRIPLE;
      break;
    case STBorder.INT_WAVE:
      xwpfType = XWPFBorderType.WAVE;
      break;
    }
    return xwpfType;
  }

  /**
   * Get the STBorderType.Enum for the specified STBorder value.
   * @param borderValue Border value (e.g., "wave").
   * @return Corresponding XWPFBorderType value or null if there is no corresponding value.
   */
  private STBorder.Enum stBorderType(XWPFBorderType borderType) {

    // There's not a direct correspondence between STBorder int values
    // and XWPFBorderType so just building a switch statement.
    STBorder.Enum stBorder = null;
    switch (borderType) {
    case DOT_DASH:
      stBorder = STBorder.DOT_DASH;
      break;
    case DASH_SMALL_GAP:
      stBorder = STBorder.DASH_SMALL_GAP;
      break;
    case DASH_DOT_STROKED:
      stBorder = STBorder.DASH_DOT_STROKED;
      break;
    case DASHED:
      stBorder = STBorder.DASHED;
      break;
    case DOT_DOT_DASH:
      stBorder = STBorder.DOT_DOT_DASH;
      break;
    case DOTTED:
      stBorder = STBorder.DOTTED;
      break;
    case DOUBLE:
      stBorder = STBorder.DOUBLE;
      break;
    case DOUBLE_WAVE:
      stBorder = STBorder.DOUBLE_WAVE;
      break;
    case INSET:
      stBorder = STBorder.INSET;
      break;
    case NIL:
      stBorder = STBorder.NIL;
      break;
    case NONE:
      stBorder = STBorder.NONE;
      break;
    case OUTSET:
      stBorder = STBorder.OUTSET;
      break;
    case SINGLE:
      stBorder = STBorder.SINGLE;
      break;
    case THICK:
      stBorder = STBorder.THICK;
      break;
    case THICK_THIN_LARGE_GAP:
      stBorder = STBorder.THICK_THIN_LARGE_GAP;
      break;
    case THICK_THIN_MEDIUM_GAP:
      stBorder = STBorder.THICK_THIN_MEDIUM_GAP;
      break;
    case THICK_THIN_SMALL_GAP:
      stBorder = STBorder.THICK_THIN_SMALL_GAP;
      break;
    case THIN_THICK_LARGE_GAP:
      stBorder = STBorder.THIN_THICK_LARGE_GAP;
      break;
    case THIN_THICK_MEDIUM_GAP:
      stBorder = STBorder.THIN_THICK_MEDIUM_GAP;
      break;
    case THIN_THICK_SMALL_GAP:
      stBorder = STBorder.THIN_THICK_SMALL_GAP;
      break;
    case THIN_THICK_THIN_LARGE_GAP:
      stBorder = STBorder.THIN_THICK_THIN_LARGE_GAP;
      break;
    case THIN_THICK_THIN_MEDIUM_GAP:
      stBorder = STBorder.THIN_THICK_THIN_MEDIUM_GAP;
      break;
    case THIN_THICK_THIN_SMALL_GAP:
      stBorder = STBorder.THIN_THICK_THIN_SMALL_GAP;
      break;
    case THREE_D_EMBOSS:
      stBorder = STBorder.THREE_D_EMBOSS;
      break;
    case THREE_D_ENGRAVE:
      stBorder = STBorder.THREE_D_ENGRAVE;
      break;
    case TRIPLE:
      stBorder = STBorder.TRIPLE;
      break;
    case WAVE:
      stBorder = STBorder.WAVE;
      break;
    }
    return stBorder;
  }
}
