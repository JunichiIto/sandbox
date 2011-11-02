use strict;
use warnings;
use ExcelUtil;

my $TO_NUMBER = 0;

main(@ARGV);

sub main {
    my @args = @_;
    my $mode = $args[0];
    my $val  = $args[1];
    my $ans;
    if ( $mode == $TO_NUMBER ) {
        $ans = ExcelUtil::to_col_number($val);
    }
    else {
        $ans = ExcelUtil::to_col_string($val);
    }
    print "$ans\n";

    return 1;
}
