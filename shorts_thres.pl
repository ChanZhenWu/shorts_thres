print "\n";
print "*******************************************************************************\n";
print "  threshold of shorts extraction tool for 3070 <v1.2>\n";
print "  Author: Noon Chen\n";
print "  A Professional Tool for Test.\n";
print "  ",scalar localtime;
print "\n*******************************************************************************\n";
print "\n";

use Excel::Writer::XLSX;


print "  please specify shorts file here: ";
   $shorts=<STDIN>;
   chomp $shorts;

print $shorts;

my $bom_coverage_report = Excel::Writer::XLSX->new($shorts.'_Thres.xlsx');
my $short_thres = $bom_coverage_report-> add_worksheet('Shorts_Thres');
$short_thres-> set_column(0,1,30);		#设置列宽
$format_head = $bom_coverage_report-> add_format(bold=>1, align=>'vcenter', border=>1, size=>12, bg_color=>'lime');
$format_data = $bom_coverage_report-> add_format(align=>'left', border=>1, size=>12);

$row = 0; $col = 0;
$short_thres-> write($row, $col, 'Nodes', $format_head);
$row = 0; $col = 1;
$short_thres-> write($row, $col, 'Threshold', $format_head);


############################### shorts threshold statistic ########################################################################

print  "\n  >>> Analyzing shorts threshold ... \n";

$node = 1;

open (Thres, "< $shorts"); 
	while($nodes = <Thres>)
	{
		chomp $nodes;
		$nodes =~ s/^ +//;	   #clear head of line spacing
		if ($nodes =~ "threshold") 
			{
				$thres = substr($nodes, index($nodes,"threshold")+10);
				if ($nodes =~ "\!"){$thres = substr($nodes, 10, index($nodes,"\!")-10);}
				$thres =~ s/\s//g;                     #clear all spacing
			}
		if ($nodes =~ "nodes")
		{
			if(substr($nodes,0,1) eq "!"){
			$short_thres-> write($node, 0, substr($nodes, 0, rindex($nodes,"!")), $format_data);  ## Nodes ##
			$short_thres-> write($node, 1, substr($nodes, rindex($nodes,"!")), $format_data);  ## Thres ##
				}
			elsif(substr($nodes,0,5) eq "nodes"){
			$short_thres-> write($node, 0, $nodes, $format_data);  ## Nodes ##
			$short_thres-> write($node, 1, $thres, $format_data);  ## Thres ##
				}
			$node++;
			#print $nodes."\n";
			}
		}
close Thres;
print "\n  done ...\n";

########################################################@##########################################################################
