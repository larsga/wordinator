package org.wordinator.xml2docx.generator;

import java.awt.image.BufferedImage;
import java.io.ByteArrayInputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.InputStream;
import java.io.IOException;
import java.net.MalformedURLException;
import java.net.URI;
import java.net.URISyntaxException;
import java.net.URL;
import java.net.URLConnection;
import javax.imageio.ImageIO;

import org.apache.commons.codec.binary.Base64;
import org.apache.commons.io.FilenameUtils;
import org.apache.commons.io.IOUtils;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.xmlbeans.XmlCursor;
import org.xml.sax.Attributes;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;
import org.xml.sax.helpers.XMLReaderFactory;
import org.xml.sax.helpers.DefaultHandler;

/**
 * Represents an image in a Word file, in order to collect all the
 * logic and data in one place.
 */
public class Image {
  private static final Logger log = LogManager.getLogger(DocxGenerator.class);

  public static void makeImage(XWPFParagraph para, XmlCursor cursor,
                               int imageCounter, File inFile, int dotsPerInch)
    throws DocxGenerationException {

    String imageUrl = cursor.getAttributeText(DocxConstants.QNAME_SRC_ATT);
    if (null == imageUrl) {
      throw new InvalidInputException("No @src attribute for image");
    }

    String widthVal = cursor.getAttributeText(DocxConstants.QNAME_WIDTH_ATT);
    String heightVal = cursor.getAttributeText(DocxConstants.QNAME_HEIGHT_ATT);
    Image image = new Image(imageUrl, imageCounter, inFile,
                            widthVal, heightVal, dotsPerInch);
    if (!image.isLoaded()) {
      return; // giving up on adding this image
    }

    XWPFRun run = para.createRun();

    try {
      run.addPicture(new ByteArrayInputStream(image.getImageBytes()),
                     image.getFormat(),
                     image.getFilename(),
                     Units.toEMU(image.getWidth()),
                     Units.toEMU(image.getHeight()));
    } catch (Exception e) {
      throw new DocxGenerationException("Exception adding picture for reference '" + image.getFilename() +"': " +
               e.getMessage());
    }
  }

  /**
   * Get the bytes from a data: URL
   * @param uri The URI that is the data: URL
   * @return The data bytes.
   */
  private static byte[] getStreamForDataUrl(URI uri) throws InvalidInputException {
     String url = uri.toString();
     // Data URL is data:[{mimeType}][;base64],{data}
     String[] tokens = url.substring(5).split(",");
     String data = tokens[1];
     String props = tokens[0];
     if (!props.matches(".*base64")) {
       throw new InvalidInputException("data: URL does not specify \"base64\", cannot decode it. URL starts with: \"" + url.substring(0, 10));
     }

     return Base64.decodeBase64(data);
  }

  /**
   * Get the MIME type from a data URL.
   * @param uri The URI that is a data URL
   * @return The MIME type as a string, or null if there is no
   * specified MIME type.
   */
  private static String getMimeTypeForDataUrl(URI uri) {
    // Data URL is data:[{mimeType}][;base64],{data}
    String[] tokens = uri.toString().substring(5).split(",");
    String props = tokens[0];
    if (props.contains(";")) {
      return props.split(";")[0];
    } else if (!"".equals(props)) {
      return props;
    }
    return null;
  }

  /**
   * Get the Word-specific format value.
   * @return The format or 0 (zero) if the format is not recognized.
   */
  private static int getImageFormat(String imgExtension) {
    if ("emf".equals(imgExtension)) return XWPFDocument.PICTURE_TYPE_EMF;
    else if ("wmf".equals(imgExtension)) return XWPFDocument.PICTURE_TYPE_WMF;
    else if ("pict".equals(imgExtension)) return XWPFDocument.PICTURE_TYPE_PICT;
    else if ("jpeg".equals(imgExtension) ||
             "jpg".equals(imgExtension)) return XWPFDocument.PICTURE_TYPE_JPEG;
    else if ("png".equals(imgExtension)) return XWPFDocument.PICTURE_TYPE_PNG;
    else if ("dib".equals(imgExtension)) return XWPFDocument.PICTURE_TYPE_DIB;
    else if ("gif".equals(imgExtension)) return XWPFDocument.PICTURE_TYPE_GIF;
    else if ("tiff".equals(imgExtension)) return XWPFDocument.PICTURE_TYPE_TIFF;
    else if ("eps".equals(imgExtension)) return XWPFDocument.PICTURE_TYPE_EPS;
    else if ("bmp".equals(imgExtension)) return XWPFDocument.PICTURE_TYPE_BMP;
    else if ("wpg".equals(imgExtension)) return XWPFDocument.PICTURE_TYPE_WPG;
    else if ("svg".equals(imgExtension)) return XWPFDocument.PICTURE_TYPE_SVG;

    return 0;
  }

  /**
   * Get the Word image format code for the specified MIME type
   * @param mimeType The MIME type to evaluate, i.e. "image/jpeg"
   * @return The format code or zero if the MIME type is not recognized.
   */
  private static int getImageFormatForMimeType(String mimeType) {
    String formatString = mimeType.split("/")[1];
    return getImageFormat(formatString);
  }

  // --- Image data

  private String imageFilename;
  private byte[] imageBytes; // must read twice, so keeping whole thing in-mem
  private String mimeType;
  private int format; // Docx format code for image
  private int width = 200; // default (why? don't know)
  private int height = 200; // default (why? don't know)

  public Image(String imageUrl, int imageCounter, File inFile,
               String widthVal, String heightVal, int dotsPerInch)
    throws DocxGenerationException {
    interpretUrl(imageUrl, inFile);

    if (!isLoaded()) {
      // we couldn't load the image, but for some reason not failing even so
      return;
    }

    // This assumes that the URL looks like a file reference.
    if (imageFilename == null || imageFilename.equals("")) {
      imageFilename = "image_" + imageCounter;
    }

    findImageFormat();
    findImageDimensions(widthVal, heightVal, dotsPerInch);
  }

  private byte[] getImageBytes() {
    return imageBytes;
  }

  public String getFilename() {
    return imageFilename;
  }

  public int getFormat() {
    return format;
  }

  public int getWidth() {
    return width;
  }

  public int getHeight() {
    return height;
  }

  public boolean isLoaded() {
    return imageBytes != null;
  }

  private void interpretUrl(String imageUrl, File inFile)
    throws DocxGenerationException {
    // Issue 72: Accept local file URLs, external URLs, and data:image/... URLs
    URI uri;
    try {
      uri = new URI(imageUrl);
    } catch (URISyntaxException e) {
      throw new InvalidInputException("Invalid URI in img/@src value: " + e.getMessage());
    }

    try {
      if (uri.isAbsolute()) {
        if ("data".equals(uri.getScheme())) {
          this.imageBytes = getStreamForDataUrl(uri);
          this.mimeType = getMimeTypeForDataUrl(uri);
        } else {
          // Should be a normal URL
          try {
            this.imageBytes = readImageData(
              uri.toURL().openConnection().getInputStream()
            );
          } catch (IOException e) {
            // it feels wrong to let this slide, but assumption that it should
            // is baked deeply into tests -- leaving it for now
            log.error("Error reading image URL: " + e.getMessage());
            return;
          }
          this.imageFilename = new File(uri.toURL().getFile()).getName();
        }

      } else {
        // Must be relative file reference, read the file
        URL baseUrl = inFile.getParentFile().toURI().toURL();
        File file = new File(new URL(baseUrl, imageUrl).getFile());
        this.imageFilename = file.getName();
        try {
          this.imageBytes = readImageData(new FileInputStream(file));
        } catch (FileNotFoundException e) {
          throw new InvalidInputException("Image file \"" + imageUrl + "\" not found.");
        }
      }
    } catch (MalformedURLException e) {
      throw new InvalidInputException("Invalid img/@src value: " + e.getMessage());
    }
  }

  private static byte[] readImageData(InputStream in)
    throws DocxGenerationException {
    try {
      byte[] bytes = IOUtils.toByteArray(in);
      in.close();
      return bytes;
    } catch (IOException e) {
      throw new DocxGenerationException("Error reading image input stream: " + e.getMessage());
    }
  }

  private void findImageFormat() throws InvalidInputException {
    String imgExtension = FilenameUtils.getExtension(imageFilename).toLowerCase();
    if (null != imgExtension && !"".equals(imgExtension)) {
      this.format = getImageFormat(imgExtension);
    } else {
      this.format = getImageFormatForMimeType(mimeType);
    }

    if (format == 0) {
        throw new InvalidInputException("Unsupported picture, format code \"" + format + "\": " + imageFilename +
                ". Expected emf|wmf|pict|jpeg|jpg|png|dib|gif|tiff|eps|bmp|wpg|svg");
    }
  }

  private void findImageDimensions(String widthVal, String heightVal,
                                   int dotsPerInch) {
    Dimensions dims = extractImageDimensions(imageFilename, imageBytes);
    int intrinsicWidth = (int) dims.getWidth();
    int intrinsicHeight = (int) dims.getHeight();

    boolean goodWidth = false;
    boolean goodHeight = false;

    // Issue 82: Handle empty width and height attributes (width="", height="")
    if (null != widthVal && !"".equals(widthVal.trim())) {
      try {
        width = (int) Measurement.toPixels(widthVal, dotsPerInch);
        goodWidth = true;
      } catch (MeasurementException e) {
        log.warn("Bad image width: " + e.getMessage());
        log.warn("Using default width value " + width);
        width = intrinsicWidth > 0 ? intrinsicWidth : width;
      }
    } else {
      width = intrinsicWidth > 0 ? intrinsicWidth : width;
    }

    if (null != heightVal && !"".equals(heightVal.trim())) {
      try {
        height = (int) Measurement.toPixels(heightVal, dotsPerInch);
        goodHeight = true;
      } catch (MeasurementException e) {
        log.warn("Bad image height: " + e.getMessage());
        log.warn("Using default height value " + height);
        height = intrinsicHeight > 0 ? intrinsicHeight : height;
      }
    } else {
      height = intrinsicHeight > 0 ? intrinsicHeight : height;
    }

    // Issue 16: If either dimension is not specified, scale the intrinsic width
    //           proportionally.
    if (widthVal == null && heightVal != null && (intrinsicWidth > 0) && goodHeight) {
      double factor = height / intrinsicHeight;
      width = (int)Math.round(intrinsicWidth * factor);
    }
    if (widthVal != null && heightVal == null && (intrinsicHeight > 0) && goodWidth) {
      double factor = (double)width / intrinsicWidth;
      height = (int)Math.round(intrinsicHeight * factor);
    }

    // At this point, the measurement is pixels. If the original specification
    // was also pixels, we need to convert to inches and then back to pixels
    // in order to apply the dots-per-inch value.

    // Word uses a DPI of 72, so if the current dotsPerInch is not 72, we need to
    // adjust the width and height by the difference.

    if (dotsPerInch != 72) {
      double factor = 72.0 / dotsPerInch;
      if (widthVal != null && widthVal.matches("[0-9]+(px)?")) {
        width =  (int)Math.round(width * factor);
      }
      if (heightVal != null && heightVal.matches("[0-9]+(px)?")) {
        height = (int)Math.round(height * factor);
      }
    }
  }

  private static Dimensions extractImageDimensions(String imageFilename,
                                                   byte[] imageBytes) {
    try {
      // FIXME: Need to limit this to the formats Java2D can read.
      BufferedImage img = ImageIO.read(new ByteArrayInputStream(imageBytes));
      if (img != null) {
        return new Dimensions(img.getWidth(), img.getHeight());
      } else {
        // let's see if we can parse it as SVG

        SvgHandler handler = new SvgHandler();
        XMLReader parser = XMLReaderFactory.createXMLReader();
        parser.setContentHandler(handler);
        parser.parse(new InputSource(new ByteArrayInputStream(imageBytes)));
        return handler.getDimensions();
      }
    } catch (IOException | SAXException e) {
      log.warn("Exception loading image file '" + imageFilename +"': " +
               e.getMessage());
      return null;
    }
  }

  private static class SvgHandler extends DefaultHandler {
    private static String SVG_NS = "http://www.w3.org/2000/svg";
    private Dimensions dimensions;

    public Dimensions getDimensions() {
      return dimensions;
    }

    @Override
    public void startElement(String uri, String localName, String qName,
                             Attributes attributes) {
      if (uri.equals(SVG_NS) && localName.equals("svg")) {
        try {
          double width = Measurement.toPixels(attributes.getValue("width"),
                                              Measurement.POINTS_PER_INCH);
          double height = Measurement.toPixels(attributes.getValue("height"),
                                               Measurement.POINTS_PER_INCH);
          this.dimensions = new Dimensions(width, height);
        } catch (MeasurementException e) {
          throw new RuntimeException(e);
        }
      }
    }
  }

  /**
   * Representing a width/height pair.
   */
  private static class Dimensions {
    private double width;
    private double height;

    public Dimensions(double width, double height) {
      this.width = width;
      this.height = height;
    }

    public double getWidth() {
      return this.width;
    }

    public double getHeight() {
      return this.height;
    }
  }
}
