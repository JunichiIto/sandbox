package ExcelUtil;
use strict;
use warnings;
use base qw( Exporter );
our @EXPORT_OK = qw( to_col_number to_col_string );

my $AZ_LENGTH  = ord('Z') - ord('A') + 1;    #26
my $OFFSET_NUM = ord('A') - 1;               #64

sub to_col_number {
    my ($col_str) = @_;
    my @chars = split //sm, $col_str;
    return convert_str( 0, \@chars );
}

sub convert_str {
    my ( $ret, $chars_ref ) = @_;

    if ( scalar( @{$chars_ref} ) == 0 ) {
        return $ret;
    }

    my $c            = shift @{$chars_ref};
    my $chars_length = scalar @{$chars_ref};
    return convert_str( calc_decimal( $c, $chars_length ) + $ret, $chars_ref );
}

sub calc_decimal {
    my ( $c, $times ) = @_;
    return $AZ_LENGTH**$times * ( ord($c) - $OFFSET_NUM );
}

sub to_col_string {
    my ($n) = @_;
    return convert_num( q{}, $n );
}

sub convert_num {
    my ( $ret, $n ) = @_;
    if ( $n == 0 ) {
        return $ret;
    }
    else {
        my ( $quo, $rem ) = excel_divmod($n);
        return convert_num( chr( $rem + $OFFSET_NUM ) . $ret, $quo );
    }
}

sub excel_divmod {
    my ($n) = @_;
    my ( $quo, $rem ) = ( int( $n / $AZ_LENGTH ), $n % $AZ_LENGTH );
    if ( $rem == 0 ) {
        return $quo - 1, $AZ_LENGTH;
    }
    else {
        return $quo, $rem;
    }
}

1;
