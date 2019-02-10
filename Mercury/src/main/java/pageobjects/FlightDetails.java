package pageobjects;

public interface FlightDetails {
	
	String ROUNDTRIP = "//*[@name='tripType'][1]";
	String ONEWAY = "//*[@name='tripType'][2]";
	String PASSENGER = "//*[@name='passCount']";
	String DEPARTINGFROM = "//*[@name='fromPort']";

}
