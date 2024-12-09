print "\n";
print "*******************************************************************************\n";
print "  threshold of shorts extraction tool for 3070 <v1.3>\n";
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
$short_thres-> set_column(0,2,30);		#设置列宽
$format_head = $bom_coverage_report-> add_format(bold=>1, align=>'vcenter', border=>1, size=>12, bg_color=>'lime');
$format_data = $bom_coverage_report-> add_format(align=>'left', border=>1, size=>12);

$row = 0; $col = 0;
$short_thres-> write($row, $col, 'Nodes', $format_head);
$row = 0; $col = 1;
$short_thres-> write($row, $col, 'Threshold', $format_head);
$row = 0; $col = 2;
$short_thres-> write($row, $col, 'Delay', $format_head);


############################### shorts threshold statistic ########################################################################

print  "\n  >>> Analyzing shorts threshold ... \n";

$node = 1;

open (Thres, "< $shorts"); 
	while($nodes = <Thres>)
	{
		chomp $nodes;
		$nodes =~ s/^ +//;	   #clear head of line spacing
		if (substr($nodes,0,9) =~ "threshold") 
			{
				$thres = substr($nodes, index($nodes,"threshold")+10);
				if ($nodes =~ "\!"){$thres = substr($nodes, 10, index($nodes,"\!")-10);}
				$thres =~ s/(^\s+|\s+$)//g;                     #clear all spacing
			}
		if ($nodes =~ "delay") 
			{
				$delay = substr($nodes, index($nodes,"delay")+6);
				#if ($nodes =~ "\!"){$delay = substr($nodes, 10, index($nodes,"\!")-10);}
				$delay =~ s/(^\s+|\s+$)//g;                     #clear all spacing
			}
		if ($nodes =~ "nodes")
		{
			if(substr($nodes,0,1) eq "!"){
			$nodes =~ s/(^\s+|\s+$)//g;                     #clear all spacing
			$node_name = substr($nodes, 0, rindex($nodes,"!"));
			$node_name =~ s/(^\s+|\s+$)//g;                     #clear all spacing
			$short_thres-> write($node, 0, $node_name, $format_data);  ## Nodes ##
			$short_thres-> write($node, 1, substr($nodes, rindex($nodes,"!")), $format_data);  ## Thres ##
			$short_thres-> write($node, 2, "-", $format_data);  ## Delay ##
				}
			elsif(substr($nodes,0,5) eq "nodes"){
			$short_thres-> write($node, 0, $nodes, $format_data);  ## Nodes ##
			$short_thres-> write($node, 1, $thres, $format_data);  ## Thres ##
			$short_thres-> write($node, 2, $delay, $format_data);  ## Delay ##
				}
			$node++;
			#print $nodes."\n";
			}
		}
close Thres;
print "\n  done ...\n";

########################################################@##########################################################################
