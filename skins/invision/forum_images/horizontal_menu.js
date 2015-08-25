// JavaScript Document

with(menuStyle){
	
	padding=5; 	//Sets the Menu Item Padding for any global menu style. 
				//Padding is the amount of white space between the text or other object and the meun item border. 
				//Padding values are in pixels. 
						
	fontsize="75%"; 	//Sets the Font Size for any global menu style. 
						//Font Sizes can be declared as any valid CSS1 or CSS2 font sizes. Default values are in pixel but em, 
						//pt etc can be used providing the value is enclosed in quotes. 
	
	fontstyle="normal"; 	//Sets the Font Style for any global menu style. Valid font styles are normal or italic 

	fontfamily="Verdana, Arial, Helvetica, sans-serif";		//Sets the Font Family for any global menu style. 
															//The Font must be present on the local computer otherwise the default browser font will be used. 
															//Font Family values can be separated by a comma with a preference to use the first font found. 
	
	pagebgcolor="#cccccc"; 	//Sets the Current Page Menu Item Background Color for any global menu style. 
							//This property, when declared, will change the background color of a menuitem if the URL matches the current URL. 
							//This is useful for indicating to the user where in the menu they need to go to navigate to this page 
	
	subimagepadding="2";	//Sets the Sub Menu Indicator Padding for any global menu style. 
							//This property will add white space to sub menu indicator images. 
							//The value for this property is in pixels. 
	
	
	onbgcolor="#A8CDF9";	//Sets the Mouse On Background Color for any global menu style. 
							//This property declares the font color for each menu item in the Mouse Off state. 
							//Valid HTML colornames or Hex values can be used for this property. 
							//When using a Hex value, the value must be prefixed by a hash. 
	
	oncolor="#465584";	//Sets the Mouse On Font Color for any global menu style. 
						//This property declares the font color for each menu item in the Mouse On state. 
						//Valid HTML colornames or Hex values can be used for this property. 
						//When using a Hex value, the value must be prefixed by a hash. 
	
	offbgcolor="#EEF2F7"; 	//Sets the Mouse Off background Color for any global menu style. 
							//All menus that use this style will use the declared background color for each menu item in the mouse of state.
							//This color should be used to declare the actual color of the menu. 
	
	offcolor="#465584";	//Sets the Mouse Off Font Color for any global menu style. 
						//This property declares the background color for each menu item in the Mouse Off state. 
						//Valid HTML colornames or Hex values can be used for this property. 
						//When using a Hex value, the value must be prefixed by a hash. 
	
	bordercolor="#006699";	//Sets the Border Color for any global Style. All valid HTML colornames or hex values can be used here. 
												//Note that the hash is needed for all hex value definitions. 
	
	borderstyle="solid";	//Sets the Border Style for any global Menu Style. 
							//All valid CSS 1 and CSS 2 values can be declared for stylename.borderstyle. 
							//Examples are: solid, dashed & dotted. 
	
	borderwidth=1;	//Sets the Border Width for any global menu style. 
					//Values are declared as pixels unless otherwise specified. 
	
	separatorcolor="#006699"; //Sets the Menuitem Separator Bar Color for any global menu style. 
												//This property sets the color of the default menu separator line in both horizontal and vertical menus. 
												//This property can also be changed inside each menu item. 
	
	separatorsize="1"; // Sets the Menuitem Separator Bar Size for any global menu style.
						//Separator size is the thickness of the separator line in pixels. 
						//used for both vertical and horizontal menus. 
}