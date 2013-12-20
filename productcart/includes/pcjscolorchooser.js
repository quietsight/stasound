<!--
// -----------------------------------------------------------------------------
// PhpConcept JavaScript - Color Chooser - PcjsColorChooser 0.2
// -----------------------------------------------------------------------------
// License GNU/GPL - Vincent Blavet - December 2005
// http://www.phpconcept.net/pcjscolorchooser
// -----------------------------------------------------------------------------
// Overview :
// Demonstration available at http://www.phpconcept.net
//
// User Guide :
//   1. Insert the script pcjscolorchooser.js in your page.
//
//   2. Call the Color Chooser by the JavaScript function
//      PcjsColorChooser(object_src_id, object_dest_id, property)
//      Where : 'object_src_id' is a valid id of the calling object
//              'object_dest_id' is a valid id of the object that will 
//              get the result
//              'property' a valid property for the destination object.
//              Can be omitted. By default will be 'value'
//
//   3. Example 1 : Selection as a input/text value :
//      <form name="form2">
//        Color :<input id='result_id' type="text" name="colortext">
//        <input id="bt1" type="button "value="Choose" 
//               onClick="PcjsColorChooser('bt1', 'result_id')">
//      </form>
//
//   4. Example 2 : Selection as a background table cell :
//      <form name="form2">
//        <table><tr><td>Color : </td>
//        <td width="20" id="testcell">&nbsp;</td>
//        <td>
//          <input id="bt2" type="button" name="Submit2" value="Choisir"
//                 onClick="PcjsColorChooser('bt2', 'testcell', 'bgColor')">
//        </td>
//        </tr></table>
//      </form>
// -----------------------------------------------------------------------------
// CVS : $Id$
// -----------------------------------------------------------------------------


// ----- Specific global values
var PcjsDestObject;
var PcjsDestProperty;

// ----- Start color chooser
function PcjsColorChooser(clickobjectid, destobjectid, destproperty)
{
  destobject = document.getElementById(destobjectid);
  
  // ----- Store the destination object
  PcjsDestObject = destobject;
  if (destproperty == "")
    PcjsDestProperty = "value";
  else
    PcjsDestProperty = destproperty;

  // ----- Get the initial value
  eval("PcjsInternalSelectColor(PcjsDestObject."+PcjsDestProperty+")");

  // ----- Select the ilayer
  var obj = document.getElementById('PcjsColorChooserPopup');

  var v_pos = PcjsccInternalGetAbsolutePosition(document.getElementById(clickobjectid));
  var v_height = document.getElementById(clickobjectid).offsetHeight;

  // ----- Set the popup position
//  obj.style.left = document.body.scrollLeft+event.clientX;
//  obj.style.top  = document.body.scrollTop+event.clientY;
  obj.style.left = v_pos.left;
  obj.style.top  = v_pos.top+v_height;

  // ----- Close the ilayer
  obj.style.visibility = "visible";
}

// -----------------------------------------------------------------------------
// Function : PcjsccInternalGetAbsolutePosition()
// Description :
//   Get the absolute position of an object.
//   The function calculates the absolute position by adding all the relative
//   offset, from the object to the oldest parent.
// -----------------------------------------------------------------------------
function PcjsccInternalGetAbsolutePosition(p_object)
{
  var v_left;
  var v_top;
  var v_position = {left:0,top:0};

  // ----- Get the object relative position
  v_position.left = p_object.offsetLeft;
  v_position.top = p_object.offsetTop;

  // ----- Get the parent absolute position
  if (p_object.offsetParent != null) {
    var v_parent_position = {left:0,top:0};
    v_parent_position=PcjsccInternalGetAbsolutePosition(p_object.offsetParent);
    v_position.left += v_parent_position.left;
    v_position.top += v_parent_position.top;
  }

  return v_position;
}
// -----------------------------------------------------------------------------

// ----- Close function for color chooser without selection
function PcjsInternalClosePopup()
{
  // ----- Select the ilayer
  var obj = document.getElementById('PcjsColorChooserPopup');

  // ----- Close the ilayer
  obj.style.visibility = "hidden";
}

// ----- Close function for color chooser with selection
function PcjsInternalSelectClose()
{
  // ----- Select the ilayer
  var obj = document.getElementById('PcjsColorChooserPopup');

  // ----- Get the value and paste it to destination object
  PcjsDestObject.value = document.forms.pcjsform.color.value;

  // ----- Look for object type
  eval("PcjsDestObject."+PcjsDestProperty+" = document.forms.pcjsform.color.value");

  // ----- Close the ilayer
  obj.style.visibility = "hidden";
}

// ----- Internal color selection
function PcjsInternalSelectColor(color)
{
  // ----- Paste the color value
  //document.forms.pcjsform.color.value = color;
  document.getElementById('pcjscolorchooser_color').value = color;

  // ----- Change the color viewer
  //document.all.pcjscolorchooser_cell.bgColor = color;
  document.getElementById('pcjscolorchooser_cell').bgColor = color;
}

// ----- Popup window creator
function PcjsGeneratePopup()
{
  // ----- Generate the div tag
  document.write("<div id=PcjsColorChooserPopup style='position:absolute; left:118px; top:214px; width:212px; height:112px; z-index:1; visibility:hidden; background-color: #FFFFFF; border: 1px none #000000'> ");
  document.write("<table width=100% border=0 cellspacing=0 bgcolor=#0000CC>");
  document.write("<tr><td height=13 width=5></td><td height=13> ");
  document.write("<div align=right><font face='Verdana, Arial, Helvetica, sans-serif' color=#FFFFFF size=2><b><a onClick='PcjsInternalClosePopup()'>x</a></b></font></div>");
  document.write("</td><td height=13 width=5> </td></tr>");
  document.write("<tr> <td height=71 width=5></td><td height=71 bgcolor=#FFFFFF> ");
  document.write("<form id=pcjscolorchooserform name=pcjsform method=post>");
  document.write("<table border=1 cellspacing=0 align=center bordercolor=#FFFFFF>");
  document.write("<tr><td align=center colspan=18><font face='Verdana, Arial, Helvetica, sans-serif' size=2><b><font color=#0000CC size=3>");
  document.write("Color Chooser</font><font color=#0000CC> </font></b></font><br>");
  document.write("</td></tr>");

  for (i=0;i<6;i++)
  {
    document.write("<tr>");
    if (i==0) v_color_i="00";
    if (i==1) v_color_i="33";
    if (i==2) v_color_i="66";
    if (i==3) v_color_i="99";
    if (i==4) v_color_i="CC";
    if (i==5) v_color_i="FF";
    for (j=0;j<6;j++)
    {
      if (j==0) v_color_j="00";
      if (j==1) v_color_j="33";
      if (j==2) v_color_j="66";
      if (j==3) {v_color_j="99"; document.write("<tr>");}
      if (j==4) v_color_j="CC";
      if (j==5) v_color_j="FF";
      for (k=0;k<6;k++)
      {
        if (k==0) v_color_k="00";
        if (k==1) v_color_k="33";
        if (k==2) v_color_k="66";
        if (k==3) v_color_k="99";
        if (k==4) v_color_k="CC";
        if (k==5) v_color_k="FF";
        document.write("<td bgcolor=#"+v_color_i+v_color_j+v_color_k+" onClick=PcjsInternalSelectColor('#"+v_color_i+v_color_j+v_color_k+"') width=10 height=10></td>");
      }
      if (j==5) {v_color_j="99"; document.write("<tr>");}
    }
    document.write("</tr>");
  }

  // ----- Basic color selection
  document.write("<tr><td colspan=18></td><td>");
  document.write("<tr><td colspan=3></td>");
  document.write("<td bgcolor=#000000 onClick=PcjsInternalSelectColor('#000000') width=10 height=10></td>");
  document.write("<td bgcolor=#333333 onClick=PcjsInternalSelectColor('#333333') width=10 height=10></td>");
  document.write("<td bgcolor=#666666 onClick=PcjsInternalSelectColor('#666666') width=10 height=10></td>");
  document.write("<td bgcolor=#999999 onClick=PcjsInternalSelectColor('#999999') width=10 height=10></td>");
  document.write("<td bgcolor=#CCCCCC onClick=PcjsInternalSelectColor('#CCCCCC') width=10 height=10></td>");
  document.write("<td bgcolor=#FFFFFF onClick=PcjsInternalSelectColor('#FFFFFF') width=10 height=10></td>");
  document.write("<td bgcolor=#FF0000 onClick=PcjsInternalSelectColor('#FF0000') width=10 height=10></td>");
  document.write("<td bgcolor=#00FF00 onClick=PcjsInternalSelectColor('#00FF00') width=10 height=10></td>");
  document.write("<td bgcolor=#0000FF onClick=PcjsInternalSelectColor('#0000FF') width=10 height=10></td>");
  document.write("<td bgcolor=#FFFF00 onClick=PcjsInternalSelectColor('#FFFF00') width=10 height=10></td>");
  document.write("<td bgcolor=#00FFFF onClick=PcjsInternalSelectColor('#00FFFF') width=10 height=10></td>");
  document.write("<td bgcolor=#FF00FF onClick=PcjsInternalSelectColor('#FF00FF') width=10 height=10></td>");
  document.write("<td colspan=3></td></tr>");

  document.write("<tr><td colspan=12 align=center>");
  document.write("<font face='Verdana, Arial, Helvetica, sans-serif' size=2 color=#0000CC>Color : </font>");
  document.write("<input type=text id=pcjscolorchooser_color name=color size=8 maxlength=8></td>");
  document.write("<td colspan=6 align=center valign=middle> ");
  document.write("<table name=tableau border=0 cellspacing=0 align=center >");
  document.write("<tr>");
  document.write("<td id=pcjscolorchooser_cell width=20 height=20 align=center valign=middle>&nbsp;</td>");
  document.write("</tr>");
  document.write("</table>");
  document.write("</td></tr>");
  document.write("<tr><td colspan=18 align=center>");
  document.write("<input type=button name=select value=Select onClick='PcjsInternalSelectClose()'>");
  document.write("</td></tr></table>");


  document.write("</form></td><td height=71 width=5></td></tr>");
  document.write("<tr height=5><td height=5 width=5></td><td height=5></td><td height=5 width=5></td></tr>");
  document.write("</table></div>");
}

// ----- Call the Color Chooser Popup Window generator function
PcjsGeneratePopup();

-->
