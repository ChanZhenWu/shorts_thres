print "\n";
print "*******************************************************************************\n";
print "  threshold of shorts extraction tool for 3070 <v1.4>\n";
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
$short_thres-> set_column(0,0,50);		#设置列宽
$short_thres-> set_column(1,3,20);		#设置列宽
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

@test_nodes = ();
@skip_nodes = ();

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
				$nodes =~ s/( +)/ /g;
# 				print $nodes,"\n";
				if($nodes =~ "\!"){$delay = substr($nodes, 15, index($nodes,"\!")-15);}
				else{$delay = substr($nodes, 15);}
# 				print $delay,"\n";
				#if ($nodes =~ "\!"){$delay = substr($nodes, 10, index($nodes,"\!")-10);}
				$delay =~ s/(^\s+|\s+$)//g;                     #clear all spacing
			}
		if ($nodes =~ "nodes")
		{
			if(substr($nodes,0,1) eq "!"){
			$nodes =~ s/(^\s+|\s+$)//g;                     #clear all spacing
			$node_name = substr($nodes, 0, rindex($nodes,"!"));
			$node_name =~ s/(^\s+|\s+$)//g;                     #clear all spacing
			#$short_thres-> write($node, 0, $node_name, $format_data);  ## Nodes ##
			#$short_thres-> write($node, 1, substr($nodes, rindex($nodes,"!")), $format_data);  ## Thres ##
			#$short_thres-> write($node, 2, "-", $format_data);  ## Delay ##
			push (@skip_nodes, $node_name."/".substr($nodes, rindex($nodes,"!"))."/"."-")
				}
			elsif(substr($nodes,0,5) eq "nodes"){
			#$short_thres-> write($node, 0, $nodes, $format_data);  ## Nodes ##
			#$short_thres-> write($node, 1, $thres, $format_data);  ## Thres ##
			#$short_thres-> write($node, 2, $delay, $format_data);  ## Delay ##
			push (@test_nodes, $nodes."/".$thres."/".$delay)
				}
			#$node++;
			#print $nodes."\n";
			}
		}
close Thres;
print "\n  done ...\n";

# @test_nodes = sort @test_nodes,"\n";
# @skip_nodes = sort @skip_nodes,"\n";

foreach my $i (0..@test_nodes-1)
{
# 	print $test_nodes[$i];
	@test_item = split("\/", $test_nodes[$i]);
	$short_thres-> write($node, 0, $test_item[0], $format_data);  ## Nodes ##
	$short_thres-> write($node, 1, $test_item[1], $format_data);  ## Thres ##
	$short_thres-> write($node, 2, $test_item[2], $format_data);  ## Delay ##
	$node++;
	}
	
foreach my $i (0..@skip_nodes-1)
{
# 	print $skip_nodes[$i];
	@test_item = split("\/", $skip_nodes[$i]);
	$short_thres-> write($node, 0, $test_item[0], $format_data);  ## Nodes ##
	$short_thres-> write($node, 1, $test_item[1], $format_data);  ## Thres ##
	$short_thres-> write($node, 2, $test_item[2], $format_data);  ## Delay ##
	$node++;
	}

$short_thres-> write(0, 3, "tested nodes: ".scalar@test_nodes, $format_anno);

########################################################@##########################################################################
