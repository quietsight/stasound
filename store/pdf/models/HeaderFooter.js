this.Header=function Header()
{
this.SetFont('Arial','',22);
this.SetTextColor(204,204,204);
this.RotatedText(7,47,'Packing Slip',90);
this.SetTextColor(0,0,0);
}
this.Footer=function Footer()
{
this.SetY(-10);
this.SetFont('Arial','',8);
this.Cell(0,10,'Diamond Store: Packing Slip ID#1431 (Code: QNU7354387902) - Page '+ this.PageNo()+ '/{nb}',0,0,'L');
}

