package oops;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.util.IOUtils;
import org.apache.poi.xslf.usermodel.SlideLayout;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFPictureData;
import org.apache.poi.xslf.usermodel.XSLFPictureShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFSlideLayout;
import org.apache.poi.xslf.usermodel.XSLFSlideMaster;
import org.apache.poi.xslf.usermodel.XSLFTextShape;

public class Hello {
	
   
	
	
   public static void main(String args[]) throws IOException{
   
      //creating a presentation 
      XMLSlideShow ppt = new XMLSlideShow();
      
      
XSLFSlideMaster slideMaster = ppt.getSlideMasters()[0];
      
      //get the desired slide layout 
      XSLFSlideLayout titleLayout = slideMaster.getLayout(SlideLayout.TITLE);
      
      //creating a slide in it 
      
      XSLFSlide slide = ppt.createSlide(titleLayout);
      XSLFSlide slide2 = ppt.createSlide();
      XSLFSlide slide3 = ppt.createSlide();
      
      
 //XSLFSlideMaster slideMaster = ppt.getSlideMasters()[0];
      
      //get the desired slide layout 
   //   XSLFSlideLayout titleLayout = slideMaster.getLayout(SlideLayout.TITLE);
                                                     
      //creating a slide with title layout
      
    ////  XSLFSlide slide1 = ppt.createSlide(titleLayout);
      
      //selecting the place holder in it 
      XSLFTextShape title1 = slide.getPlaceholder(0); 
      
      //setting the title init 
      title1.setText("OOPs What is happening here");
      //reading an image
      
      
      File image=new File("G://images//1493239_735336023214399_7168766204430241554_n.jpg");
      
      
      File image2=new File("G://images//10440718_346149075563401_2889381141235210400_n.jpg");
      
      
      
      //converting it into a byte array
      byte[] picture = IOUtils.toByteArray(new FileInputStream(image));
      byte[] picture2=IOUtils.toByteArray(new FileInputStream(image2));
      
      //adding the image to the presentation
      int idx = ppt.addPicture(picture, XSLFPictureData.PICTURE_TYPE_PNG);
      int idx1 = ppt.addPicture(picture2, XSLFPictureData.PICTURE_TYPE_PNG);
      
      
      
      //creating a slide with given picture on it
      XSLFPictureShape pic = slide.createPicture(idx);
      XSLFPictureShape pic2 = slide2.createPicture(idx1);
      
      //XSLFPictureShape pic3 = slide3.createPicture(idx);
      
      
      
      
      
      
      
      
      
      java.awt.Dimension pgsize = ppt.getPageSize();
      int pgw = pgsize.width; //slide width in points
      int pgh = pgsize.height; //slide height in points
      System.out.println("current page size of the PPT is:");
      System.out.println("width :" + pgw);
      System.out.println("height :" + pgh);
      
      //set new page size
      
      
      ppt.setPageSize(new java.awt.Dimension(1340,1030));
      
      
      
      
      
      
      
      //creating a file object 
      File file=new File("addingimageGood7.pptx");
      
      
      
      
      FileOutputStream out = new FileOutputStream(file);
      
      //saving the changes to a file
      ppt.write(out);
      System.out.println("image added successfully");
      out.close();	
   }
}