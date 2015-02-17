package oops;


import java.awt.Graphics2D;
import java.awt.geom.Rectangle2D;
import java.awt.image.BufferedImage;
import java.io.ByteArrayInputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import javax.imageio.ImageIO;

import org.apache.poi.util.IOUtils;
import org.apache.poi.xslf.usermodel.SlideLayout;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFPictureData;
import org.apache.poi.xslf.usermodel.XSLFPictureShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFSlideLayout;
import org.apache.poi.xslf.usermodel.XSLFSlideMaster;
import org.apache.poi.xslf.usermodel.XSLFTextParagraph;
import org.apache.poi.xslf.usermodel.XSLFTextRun;
import org.apache.poi.xslf.usermodel.XSLFTextShape;



public class Exporter {
	
	
	public static void main(String args[]) throws IOException{
	   
		
	   int ppt2013Width=1368;
	   int ppt2013Height=768;
	   int pageWidth=1000;
	   int pageHeight=850;
	   String sourceFolder = "G://dinesh//Source Images";
       XMLSlideShow ppt = new XMLSlideShow();
    
       File image2=new File("G://dinesh//Images//Right.jpg");
       
       File image3=new File("G://dinesh//Images//Left.jpg");
       
       File image4=new File("G://dinesh//Source Images//15.Legend1 (legend).png");
       File image5=new File("G://dinesh//Source Images//16.Legend2 (legend).png");
    		  
       
       //converting it into a byte array
     
       byte[] picture2=IOUtils.toByteArray(new FileInputStream(image2));
       byte[] picture3=IOUtils.toByteArray(new FileInputStream(image3));
       
       byte[] picture4=IOUtils.toByteArray(new FileInputStream(image4));
       byte[] picture5=IOUtils.toByteArray(new FileInputStream(image5));
       
       //adding the image to the presentation
    
       int idx1 = ppt.addPicture(picture2, XSLFPictureData.PICTURE_TYPE_PNG);
       
       int idx2 = ppt.addPicture(picture3, XSLFPictureData.PICTURE_TYPE_PNG);
       int idx3 = ppt.addPicture(picture4, XSLFPictureData.PICTURE_TYPE_PNG);
       int idx4 = ppt.addPicture(picture5, XSLFPictureData.PICTURE_TYPE_PNG);
       
       //creating a slide with given picture on it
    
       //XSLFPictureShape pic2 = slide2.createPicture(idx1);
       
       File folder = new File(sourceFolder);
       
       File[] listOfFiles = folder.listFiles();
        
       for (File image : listOfFiles) {
           if(image.isFile() && image.getName().endsWith("png")&&(image.getName().indexOf("Legend")==-1)) {
        	           	   BufferedImage img = null;

               try {
                   img = ImageIO.read(image);
                   
                   XSLFSlideMaster slideMaster = ppt.getSlideMasters()[0];
                   
                   //get the desired slide layout 
                    XSLFSlideLayout titleLayout = slideMaster.getLayout(SlideLayout.TITLE_ONLY);
                   //   titleLayout.createAutoShape();
                  
                   
               XSLFSlide slide = ppt.createSlide(titleLayout);
               
               slide.setFollowMasterGraphics(true);
               //slide.setBackground(new java.awt.Color(0,0,255));
               
               XSLFTextShape title1 = slide.getPlaceholder(0);
               // XSLFTextShape body = slide.getPlaceholder(1);
               
               //clear the existing text in the slide
               //body.clearText();
               
               //adding new paragraph
               //body.addNewTextParagraph().addNewTextRun().setText("this is  my first slide body");
               
               
               // XSLFTextShape body = slide.getPlaceholder(1);
               
               //clear the existing text in the slide
               // body.clearText();
               
               
               int width=img.getWidth();
               int height=img.getHeight();
               // int type=img.getType();
               
             
               
               //following expression is currently unused
               if(width>pageWidth){
                   double divider = width/pageWidth;
                   double multiplier = height/width;
                   width /= divider;
                   //width = Math.floor(width);
                   //height = Math.floor(width*multiplier);
               }
               if(height>pageHeight){
                   double divider = height/width;
                  // width = Math.floor(pageHeight/divider);
                   height = pageHeight;
               }
               
               
           //int ee=img.getRGB(IMG_HEIGHT, IMG_WIDTH);
               
               
           //  BufferedImage resizedImage = new BufferedImage(IMG_WIDTH, IMG_HEIGHT, type);
             
               //adding new paragraph
               // body.addNewTextParagraph().addNewTextRun().setText("this is  my first slide body");
               
               //setting the title init 
               java.awt.Dimension pgsize = ppt.getPageSize();
               int pgw = pgsize.width; //slide width in points
               int pgh = pgsize.height; //slide height in points
               System.out.println("current page size of the PPT is:");
               System.out.println("width :" + pgw);
               System.out.println("height :" + pgh);
               
               
               String a=image.getName().replace(".png", "");
               
               if(image.getName().indexOf('(')>-1){
					a = image.getName().substring(0,image.getName().indexOf('('));
				}
               a = a.substring(a.indexOf('.')+1);
               
               title1.setText(a.replaceAll("[(1-9)]+","" ));
               title1.setAnchor(new java.awt.Rectangle(251,45,700,15));
               title1.setFillColor(java.awt.Color.green);
               title1.setWordWrap(true);
               //XSLFTextParagraph paragraph=title1.addNewTextParagraph();
               //XSLFTextRun run = paragraph.addNewTextRun();
               //run.setFontColor(java.awt.Color.red);
              // run.setFontSize(24);
                byte[] picture = IOUtils.toByteArray(new FileInputStream(image));
                
                int idx = ppt.addPicture(picture, XSLFPictureData.PICTURE_TYPE_PNG);
                
                XSLFPictureShape pic = slide.createPicture(idx);
                
                
                XSLFPictureShape pic2 = slide.createPicture(idx1);
                XSLFPictureShape pic3 = slide.createPicture(idx2);
                
                
               // String ch=image.getName();
           	 //if (ch.contains("Timeliness"))
           	 //{
           		 
           	//	XSLFPictureShape pic4 = slide.createPicture(idx3);
           	 // pic4.setAnchor(new java.awt.Rectangle(300,628,300,140));
           	 //}
           	 //else
           	 //{
           	//	XSLFPictureShape pic5 = slide.createPicture(idx4);
           	 //   pic5.setAnchor(new java.awt.Rectangle(300,628,300,140));
           	 //}
                if(image.getName().contains("map"))
					if(image.getName().contains("Timeliness")){
						XSLFPictureShape pic4 = slide.createPicture(idx3);
						pic4.setAnchor(new java.awt.Rectangle(190,500,250,140));
					}else{
						XSLFPictureShape pic5 = slide.createPicture(idx4);
						pic5.setAnchor(new java.awt.Rectangle(190,500,250,140));
					}
				
           	 
           	 
           	 if (height>(ppt2013Height-150))
           		 
           	 {
           		 if(height>width)
           		 {
           			 
           			 int heightChanged=height+100;
               		 if (heightChanged<668)
               		 {
           			 
           			pic.setAnchor(new java.awt.Rectangle((((ppt2013Width-ppt2013Width/8)-654)/2),100,654,heightChanged));
               		 }
               		 else{
               			pic.setAnchor(new java.awt.Rectangle((((ppt2013Width-ppt2013Width/8)-654)/2),100,654,668));
               		 }
           			
           			
           			
           			
           			
           		 }
           		 else{
                pic.setAnchor(new java.awt.Rectangle((((ppt2013Width-ppt2013Width/8)-704)/2),100,704,508));
           	 }}
           	// else if(width<500||height<400)
           	 //{
           		 
           		// if(width>height){
           		 
           		 //pic.setAnchor(new java.awt.Rectangle(181,100,504,400));
           		 //}
           		 //else{
           		//	pic.setAnchor(new java.awt.Rectangle(181,100,400,500));
           		 //}
           	 //}
           	 else
           	 {
           		 int heightChanged=height+100;
           		 if (heightChanged<618)
           		 {
           		 pic.setAnchor(new java.awt.Rectangle((((ppt2013Width-ppt2013Width/8)-width)/2),100,width+100,heightChanged)); 
           		 }
           		 else{
           			 pic.setAnchor(new java.awt.Rectangle((((ppt2013Width-ppt2013Width/8)-width)/2),100,width+100,height)); 
                		
           		 }	 
           	 }
           	 
                pic2.setAnchor(new java.awt.Rectangle((ppt2013Width-ppt2013Width/8),0,(ppt2013Width/8),ppt2013Height));
                
                pic3.setAnchor(new java.awt.Rectangle(0,0,(ppt2013Width/8),ppt2013Height));
                
              
                
                
               //pic.getPictureData();
               
                
                
                
                // pic.resize();
                // java.awt.Dimension imgsize = pic.get
                
                //pic.drawContent(g)
                System.out.println(image.getName());
                System.out.println("path:" + image.getPath());
                System.out.println(" width : " + img.getWidth());
                System.out.println(" height: " + img.getHeight());
                System.out.println(" size  : " + image.length());
                System.out.println("" + image.getParent());
                System.out.println("" + image.getCanonicalPath());
                System.out.println("" + image.getAbsolutePath());
               } catch (final IOException e) {
                   // handle errors here
               }
                
            }
        }       
       //set new page size
      ppt.setPageSize(new java.awt.Dimension(ppt2013Width,ppt2013Height));

      File file =new File(sourceFolder+"\\Export7.pptx");
      FileOutputStream out = new FileOutputStream(file);
      
      ppt.write(out);
      System.out.println("Presentation created successfully");
      out.close();
         }
}