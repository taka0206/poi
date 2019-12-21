import javax.xml.ws.*;

public class PoiServer {
  public static void main(String[] args) {

  	Endpoint.publish( 
       "http://localhost:8888/WebServices/PoiWrapper", 
       new PoiWrapper()
       );
  }
}
