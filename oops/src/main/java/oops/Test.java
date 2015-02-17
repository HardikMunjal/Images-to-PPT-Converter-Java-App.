package oops;

import java.awt.image.BufferedImage;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FilenameFilter;
import java.io.IOException;

import javax.imageio.ImageIO;

import org.apache.poi.util.IOUtils;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFPictureData;
import org.apache.poi.xslf.usermodel.XSLFPictureShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;


public class Test {

    // File representing the folder that you select using a FileChooser
    static final File dir = new File("G://images");

    
    // array of supported extensions (use a List if you prefer)
    static final String[] EXTENSIONS = new String[]{
        "jpg", "png", "bmp" // and other formats you need
    };
    
    // filter to identify images based on their extensions
    static final FilenameFilter IMAGE_FILTER = new FilenameFilter() {

        //@Override
        public boolean accept(final File dir, final String name) {
            for (final String ext : EXTENSIONS) {
                if (name.endsWith("." + ext)) {
                    return (true);
                }
            }
            return (false);
        }
    };

    public static void main(String[] args)throws IOException {
    	
    	
    	

        if (dir.isDirectory()) { // make sure it's a directory
            for (final File f : dir.listFiles(IMAGE_FILTER)) {
                BufferedImage img = null;

                try {
                    img = ImageIO.read(f);
                    
                    
                    
                    
                    
                    XMLSlideShow ppt = new XMLSlideShow();	     
                    
                    for(int i=0;i<10;i++)
                    {
                    XSLFSlide slide = ppt.createSlide();
                    
                    

                    File image=new File("G://images//1493239_735336023214399_7168766204430241554_n.jpg");





                    //converting it into a byte array
                    byte[] picture = IOUtils.toByteArray(new FileInputStream(image));


                    //adding the image to the presentation
                    int idx = ppt.addPicture(picture, XSLFPictureData.PICTURE_TYPE_PNG);




                    //creating a slide with given picture on it
                    XSLFPictureShape pic = slide.createPicture(idx);
                    
                    }

                    //XSLFPictureShape pic3 = slide3.createPicture(idx);




                            //creating an FileOutputStream object
                            File file =new File("example15.pptx");
                            FileOutputStream out = new FileOutputStream(file);
                        	

                    
                    
                    
                    
                    
                    
                    
                            ppt.write(out);
                            System.out.println("Presentation created successfully");
                        
                        
                    
                    

                    // you probably want something more involved here
                    // to display in your UI
                    System.out.println("image: " + f.getName());
                    System.out.println("path:" + f.getPath());
                    System.out.println(" width : " + img.getWidth());
                    System.out.println(" height: " + img.getHeight());
                    System.out.println(" size  : " + f.length());
                    System.out.println("" + f.getParent());
                    System.out.println("" + f.getCanonicalPath());
                    System.out.println("" + f.getAbsolutePath());
                    System.out.println("" + f.pathSeparator);
                    
                } catch (final IOException e) {
                    // handle errors here
                }
            }
        }
        
           
        

        
       
        
    }
}