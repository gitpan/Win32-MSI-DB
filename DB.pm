package Win32::MSI::DB;

=head1 NAME

Win32::MSI::DB - a Perl package to modify MSI databases

=head1 SYNOPSIS

  use Win32::MSI::DB;
  
  $database=Win32::MSI::DB::new('filename', $flags);

  $database->transform('filename', $flags);

  $table=$database->table("table");
  $view=$database->view("SELECT * FROM File 
			WHERE FileSize < ?",100000);
  
  @rec=$table->records();
  $rec4=$table->record(4);
  
  $rec->set("field","value"); # string
  $rec->set("field",4);       # int
  $rec->set("field","file");  # streams
  
  $rec->get("field");
  $rec->getintofile("field","file");
  
  $field=$rec->field("field");
  $field->set(2);
  $data=$field->get();
  $field->fromfile("autoexec.bat");
  $field->intofile("tmp.aa");
  
  $db->error();
  $view->error();
  $rec->error();

=head1 DESCRIPTION

=head2 Obtaining a database object

This is currently done by using C<MSI::DB::new>.
It takes F<filename> as first parameter and one of the following constants as optional second. 

=over 4

=item $Win32::MSI::MSIDBOPEN_READONLY

This opens the file B<not> read-only, but changes will not be written to disk.

=item $Win32::MSI::MSIDBOPEN_TRANSACT

This allows transactional functionality, ie. changes are written on commit only.
This is the default.

=item $Win32::MSI::MSIDBOPEN_DIRECT

Opens read/write without transactional behaviour.

=item $Win32::MSI::MSIDBOPEN_CREATE

This creates a new database in transactional mode.

=back

This database object allows creation of C<table> or C<view>s, depending on your needs.
If you simply need access to a table you can use the C<table> method; for a subset of records or even a SQL-query you can use the
C<view> method.

=head2 Using transforms

When you have got a handle to a database, you can successivly apply transforms to it. You do this by using C<transform>, which needs the filename of the transform file (normally with extension F<.mst>) and optionally a flag specification.

Most of the possible flag values specify which merge errors are to be suppressed.

=over 4

=item $Win32::MSI::MSITR_IGNORE_ADDEXISTINGROW

Ignores adding a row that already exists.

=item $Win32::MSI::MSITR_IGNORE_ADDEXISTINGTABLE

Ignores adding a table that already exists.

=item $Win32::MSI::MSITR_IGNORE_DELMISSINGROW

Ignores deleting a row that doesn't exist.

=item $Win32::MSI::MSITR_IGNORE_DELMISSINGTABLE

Ignores deleting a table that doesn't exist.

=item $Win32::MSI::MSITR_IGNORE_UPDATEMISSINGROW

Ignores updating a row that doesn't exist.

=item $Win32::MSI::MSITR_IGNORE_CHANGECODEPAGE

Ignores that the code pages in the MSI database and the transform file do not match and neither has a neutral code page.

=item $Win32::MSI::MSITR_IGNORE_ALL

This flag combines all of the above mentioned flags. This is the default.

=item $Win32::MSI::MSITR_VIEWTRANSFORM

This flag should not be used together with the other flags. It specifies that instead of merging the data a table named C<_TransformView> is created in memory, which has the columns C<Table>, C<Column>, C<Row>, C<Data> and C<Current>.

This way the data in a transform file can be directly queried.

For more information please see S<http://msdn.microsoft.com/library/default.asp?url=/library/en-us/msi/setup/_transformview_table.asp>.

=back

This opens the file B<not> read-only, but changes will not be written to disk.

A transform is a specification of changed values. So you get a MSI database from your favorite vendor, make a transform to <overlay> your own settings (the target installation directory, the features to be installed, etc.) and upon installation you can use these settings via a commandline similar to

  msiexec /i TRANSFORMS=F<your transform file> F<the msi database> /qb

The changes in a transform are stored by a (table, row, cell, old value, new value) tuple.

=head2 From a table or view to records

When you have obtained a C<table> or C<view> object, you can use the
C<record> method to access individual records. It takes a number as parameter.
Here the records are fetched as needed; using C<undef> as parameter fetches all records and returns the first (index 0).

Another possibility is to use the method C<records>, which returns an array of all records in this table or view.

=head2 A record has fields

And this fields can be queried or changed using the C<record> object, as in 

  $rec->set("field","value"); # string
  $rec->set("field",4);       # int
  $rec->set("field","file");  # streams
  
  $rec->get("field");
  $rec->getintofile("field","file");

or you can have separate C<field> objects:

  $field=$rec->field("field");

  $data=$field->get();
  $field->set(2);

Remark: the access to files (streams) is currently not finished.

=head2 Errors

Each object may access an C<error> method, which gives a string or an array (depending on context)
containing the error information.

Help wanted: Is there a way to get a error string from the number which does not depend on the current MSI database?

Especially the developer-errors (2000 and above) are not listed.

=head1 REMARKS

This module depends on C<Win32::API>, which is used to import the functions out of the F<msi.dll>.

Currently the C<Exporter> is not used - patches are welcome.

=head2 AUTHOR

Please contact C<pmarek@cpan.org> for questions, suggestions, and patches (C<diff -wu2> please).

=head2 Further plans

A C<Win32::MSI::Tools> package is planned - which will allow to compare databases and give a diff, and similar tools.

I have started to write a simple Tk visualization.

=head1 SEE ALSO

S<http://msdn.microsoft.com/library/default.asp?url=/library/en-us/msi/setup/installer_database_reference.asp>

=cut

use Win32::API;

$VERSION="1.04";


###### Constants and other definitions

$MsiOpenDataBase =new Win32::API("msi","MsiOpenDatabase","PPP","I") || die $!;
$MsiOpenDataBasePIP =new Win32::API("msi","MsiOpenDatabase","PIP","I") || die $!;
$MsiCloseHandle =new Win32::API("msi","MsiCloseHandle","I","I") || die $!;
$MsiDataBaseCommit =new Win32::API("msi","MsiDatabaseCommit","I","I") || die $!;
$MsiDatabaseApplyTransform =new Win32::API("msi","MsiDatabaseApplyTransform","IPI","I") || die $!;

$MsiViewExecute =new Win32::API("msi","MsiViewExecute","II","I") || die $!;
$MsiDatabaseOpenView =new Win32::API("msi","MsiDatabaseOpenView","IPP","I") || die $!;
$MsiViewClose =new Win32::API("msi","MsiViewClose","I","I") || die $!;
$MsiViewFetch =new Win32::API("msi","MsiViewFetch","IP","I") || die $!;

$MsiRecordGetFieldCount =new Win32::API("msi","MsiRecordGetFieldCount","I","I") || die $!;
$MsiRecordGetInteger  =new Win32::API("msi","MsiRecordGetInteger","II","I") || die $!;
$MsiRecordGetString =new Win32::API("msi","MsiRecordGetString","IIPP","I") || die $!;
$MsiRecordGetStringIIIP =new Win32::API("msi","MsiRecordGetString","IIIP","I") || die $!;

$MsiRecordSetInteger  =new Win32::API("msi","MsiRecordSetInteger","III","I") || die $!;
$MsiRecordSetString =new Win32::API("msi","MsiRecordSetString","IIP","I") || die $!;
$MsiRecordSetStream =new Win32::API("msi","MsiRecordSetStream","IIP","I") || die $!;
$MsiCreateRecord =new Win32::API("msi", "MsiCreateRecord", "I","I") || die $!;

$MsiViewGetColumnInfo =new Win32::API("msi", "MsiViewGetColumnInfo", "IIP","I") || die $!;

$MsiGetLastErrorRecord =new Win32::API("msi", "MsiGetLastErrorRecord", "", "I") || die $!;
$MsiFormatRecord =new Win32::API("msi", "MsiFormatRecord", "IIPP", "I") || die $!;


$MSIDBOPEN_READONLY=0;
$MSIDBOPEN_TRANSACT=1;
$MSIDBOPEN_DIRECT=2;
$MSIDBOPEN_CREATE=3;

$MSICOLINFO_NAMES=0;
$MSICOLINFO_TYPES=1;
$_MSICOLINFO_INDEX=21231231; # for own use, not defined by MS


$MSITR_IGNORE_ADDEXISTINGROW=0x1;
$MSITR_IGNORE_DELMISSINGROW=0x2;
$MSITR_IGNORE_ADDEXISTINGTABLE=0x4;
$MSITR_IGNORE_DELMISSINGTABLE=0x8;
$MSITR_IGNORE_UPDATEMISSINGROW=0x10;
$MSITR_IGNORE_CHANGECODEPAGE=0x20;
$MSITR_VIEWTRANSFORM=0x100;

$MSITR_IGNORE_ALL=
  $MSITR_IGNORE_ADDEXISTINGROW |
  $MSITR_IGNORE_DELMISSINGROW |
  $MSITR_IGNORE_ADDEXISTINGTABLE |
  $MSITR_IGNORE_DELMISSINGTABLE |
  $MSITR_IGNORE_UPDATEMISSINGROW |
  $MSITR_IGNORE_CHANGECODEPAGE;

$MSI_NULL_INTEGER=-0x80000000;
$ERROR_NO_MORE_ITEMS||=259;
$ERROR_MORE_DATA||=234;


$COLTYPE_STREAM = 1;
$COLTYPE_INT = 2;
$COLTYPE_STRING = 3;
%COLTYPES = (
  "i" => $COLTYPE_INT,
  "j" => $COLTYPE_INT,
  "s" => $COLTYPE_STRING,
  "g" => $COLTYPE_STRING,
  "l" => $COLTYPE_STRING,
  "v" => $COLTYPE_STREAM,
);

$INITIAL_EMPTY_STRING= "\0" x 1024;

##### Default Routines
sub new
{
  my($file,$mode)=@_;
  my(%a,$hdl);
  my($me);

  return undef unless $file;

  $hdl="\0\0\0\0";
  $mode=$MSIDBOPEN_TRANSACT if !defined $mode;
  if ($mode =~ /^\d+$/)
  {
# For special values of mode another call is 
# needed (integer instead of pointer)
    $MsiOpenDataBasePIP->Call($file, $mode,$hdl) && return undef;
  }
  else
  {
    $MsiOpenDataBase->Call($file, $mode,$hdl) && return undef;
  }

  $a{"handle"}=unpack("l",$hdl);

  _bless_type(\%a,"db");
}

sub DESTROY
{
  my $self=shift;

  $self->_commit()
  if ($self->{""} eq "db");

  &_close($self->{"handle"}) && return undef
  if $self->{"handle"};
  $self={};
}


##### External Routines

sub table
{
  my($self,$table, $where,@parm)=@_;
  my($sql);


  $sql="SELECT * FROM " . $table . "";
  $sql.=" WHERE " . $where if ($where);

  $self->view($sql,@parm);
}

sub view
{
  my($self,$sql,@parm)=@_;
  my($hdl);
  my(%s,$a,$me);

  $self->_check("db");

  $hdl="\0\0\0\0";
  $a=$MsiDatabaseOpenView->Call($self->{"handle"},$sql,$hdl);
  $a && return undef;

  $s{"handle"}=unpack("l",$hdl);
  if (@parm)
  {
    $a=_newrecord(@parm);
    $a || return undef;
#		print "openview: ",scalar(@parm)," parms: ",join(" ",@parm),"\n";
  }
  else
  {
    $a=0;
  }
  $MsiViewExecute->Call($s{"handle"},$a) && return undef;

  _close($a) if ($a);


  if ($sql !~ /^\s*SELECT\s/i)
  {
    $me=_bless_type(\%s,"sql");
    return $me;
  }

  $me=_bless_type(\%s,"v");
  $me->get_info(undef);
  $me->{"coltypes"}=
  [ map { 
    $COLTYPES{ 
      lc( 
	substr($_->{"type"},0,1) 
      ) 
    }; 
  } @{$me->{"colinfo"}} ];

  $me;
}

sub record
{
  my($self,$recnum)=@_;

  $self->_check("v");

  while ($recnum > $self->{"fetched"} || !defined($recnum))
  {
    my($hdl,$l);

    $hdl="\0\0\0\0";
    $l=$MsiViewFetch->Call($self->{"handle"},$hdl);

    last if ($l == $ERROR_NO_MORE_ITEMS);

    $hdl=unpack("l",$hdl);
    $self->{"records"}[$self->{"fetched"} ++]=
    _bless_type(
      { "handle" => $hdl, 
	"view" => $self }, 
      "r");
  }  

  $self->{"records"}[$recnum];
}

sub records
{
  my($self)=@_;

  $self->_check("v");

  $self->record(undef);
  @{$self->{"records"}};
}

sub fields
{
  return field(@_);
}

sub field
{
  my($self,@name)=@_;
  my(@ret,$n,$i,$cn);

  $self->_check("r");
  @ret=();
  for $n (@name)
  {
    $i=$self->{"view"}->get_info($_MSICOLINFO_INDEX,$field);
    if (defined $i)
    {
      push @ret,
      _bless_type(
	{ "rec" => $self,
	  "cn" => $i->{"index"} }, 
	"f");
    }
    else
    { 
      push @ret,undef;
    }
  }

  @name > 1 || wantarray() ? @ret : $ret[0];
}

sub close
{
  my $self=shift;
  $self->DESTROY();
}

sub get
{
  my($self,$field)=@_;
  my($f);

  $self->_check("r","f");

  if ($self->_type() eq "f")
  {
#field
    return $self->{"rec"}{"data"}[$self->{"cn"}];
  }

# record
  if (!$self->{"data"})
  {
    $self->{"data"} = [ 
    _extract_fields(
      $self->{"handle"},
      @{$self->{"view"}{"coltypes"}} ) ];
  }

  $f=$self->{"view"}->get_info($_MSICOLINFO_INDEX,$field);

  return defined($f) ? $self->{"data"}[$f] : undef;
}

sub set
{
  my($self,$field,$value)=@_;
  my($rec,$cn,$type);

  $self->_check("r","f");

  if ($self->_type() eq "r")
  {
# record
    $rec=$self;
    $cn=$self->{"view"}->get_info($_MSICOLINFO_INDEX,$field);
  }
  else
  {
# field
    $rec=$self->{"rec"};
    $cn=$self->{"cn"};
    $value=$field; # $field not given
  }

# msi numbers columns from 1
  $type=$rec->{"view"}{"coltypes"}[$cn];
  if ($type == $COLTYPE_INT)
  {
    $MsiRecordSetInteger->Call($rec->{"handle"},$cn+1,$value) && return undef;
  }
  elsif ($type == $COLTYPE_STRING)
  {
    $MsiRecordSetString->Call($rec->{"handle"},$cn+1,$value) && return undef;
  }
  elsif ($type == $COLTYPE_STREAM)
  {
    $MsiRecordSetStream->Call($rec->{"handle"},$cn+1,$value) && return undef;
  }
  else
  {
    return undef;
  }

  return 1;
}

sub error
{
  my($self)=shift;
  my ($e,$q);
  my(@a,$s,$l);

  $e=$MsiGetLastErrorRecord->Call();
#  die "no error" if (!$e);
  return undef if (!$e);

  @a=_extract_fields($e);
  _close($e);

  return wantarray() ? @a : "MSIDB error: '" . join("' '",@a) . "'";

# TODO: is there some way we can get the text of the error messages?
# they are only partly in the msi file. (only for installation)
#	developer error codes (wrong SQL syntax eg) ?
  print join("<>",@a),"\n";
  $q=$self->openview("SELECT Message FROM Error WHERE Error=?",$a[0]);
  die unless $q;

  push @a,$q->fetch();
  &_close($q);
  $q=newrecord(@a) || die $!;
  print "rec=$q\n";
  $s=" " x 1024;
  $l=pack("l",length($s));
  $MsiFormatRecord->Call($self,$q,$s,$l) || die $!;
  print "->$s\n";
  substr($s,unpack("l",$l))="";
  &_close($e);
  $s;
}

sub coltypes
{
  my($self)=@_;
  $self->get_info($MSICOLINFO_TYPES);
}

sub colnames
{
  my($self)=@_;
  $self->get_info($MSICOLINFO_NAMES);
}

sub get_info
{
  my($self,$which,$field)=@_;
# is MSICOLINFO_NAMES MSICOLINFO_TYPES
  my $hdl,@name,@type,$n,$t,$i;

  $self->_check("v");

  if (!$self->{"colinfo"})
  {
    $hdl="\0\0\0\0";
    $MsiViewGetColumnInfo->Call($self->{"handle"},$MSICOLINFO_NAMES,$hdl) && return undef;
    $hdl=unpack("l",$hdl);
    @name= _extract_fields($hdl);
    &_close($hdl);

    $hdl="\0\0\0\0";
    $MsiViewGetColumnInfo->Call($self->{"handle"},$MSICOLINFO_TYPES,$hdl) && return undef;
    $hdl=unpack("l",$hdl);
    @type= _extract_fields($hdl);
    &_close($hdl);

    $i=0;
    while (@name)
    {
      $n=shift @name;
      $t=shift @type;
      $self->{"colinfo_hash"}{$n} = $self->{"colinfo"}[$i] = 
      { "name" => $n, 
	"type" => $t,
	"index" => $i};
      $i++;
    }
  }

  if ($which == $_MSICOLINFO_INDEX)
  {
    return undef unless $field;

    $t=$self->{"colinfo_hash"}{$field};
    return $t ? $t->{"index"} : undef;
  }

  return %{$self->{"colinfo_hash"}} if !defined($which);
  return map { $_->{"name"}; } @{$self->{"colinfo"}} 
  if ($which == $MSICOLINFO_NAMES);
  return map { $_->{"type"}; } @{$self->{"colinfo"}} 
  if ($which == $MSICOLINFO_TYPES);
  return undef;
}

sub die
{
  my($self)=shift;
  my(@a,@c);

  @a=$self->error();
  @c=caller;
  print "error ",shift @a;
  print ": ",join(" ",@a),"\n",
  "in ",join(" ",@c),"\n";
}


sub transform
{
  my($self,$filename,$flags)=@_;
  my($r);

  $self->_check("db");
  return undef unless $filename;

  $flags=$MSITR_IGNORE_ALL if !defined($flags);

  $r=$MsiDatabaseApplyTransform->Call(
    $self->{"handle"},$filename,$flags);
  return $r;
}


##### Internal Routines
# should not be used outside this module


sub _commit
{
  my $self=shift;
  $MsiDataBaseCommit->Call($self->{"handle"}) && return undef;
}

sub _close
{
  my $hdl=shift;

  $MsiCloseHandle->Call($hdl) && return undef;
}

sub _type
{
  my($self)=@_;

  return $self->{""};
}

sub _check
{
  my($self,@allowed)=@_;
  my($t);

  $t=$self->_type();
  die "$self is a wrong type:'$t' instead of " . join(",",@allowed)
  unless grep($t eq $_,@allowed);
}

sub _bless_type
{
  my($ref,$type,$class)=@_;
  my($me);

  $me=bless $ref,$class || "Win32::MSI::DB";
  $me->{""}=$type;
  $me;
}

sub _newrecord
{
  my(@list)=@_;
  my($hdl,$i);

  $hdl=$MsiCreateRecord->Call(scalar(@list));
  return undef if !$hdl;

  for($i=0; $i<@list; $i++)
  {
#		print "new rec. $i: ",$list[$i]," is a ";
    if ($list[$i] =~ /^\d+$/)
    {
#			print "int\n";
      $MsiRecordSetInteger->Call($hdl,$i+1,$list[$i]) && return undef;
    }
    else
    {
#			print "string\n";
      $MsiRecordSetString->Call($hdl,$i+1,$list[$i]) && return undef;
    }
  }

  $hdl;
}

sub _getI
{
  my($hdl,$num)=@_;
  my($i);

  $i= $MsiRecordGetInteger->Call($hdl,$num);
  return undef if ($i == $MSI_NULL_INTEGER);
  $i;
}

sub _getS
{
  my($hdl,$num)=@_;
  my($l,$s,$e,$p);

  $s=$INITIAL_EMPTY_STRING;
  $p=pack("l",length($s)); # initial size
  $e=$MsiRecordGetString->Call($hdl,$num,$s,$p);
  if ($e == $ERROR_MORE_DATA)
  {
    $l=unpack("l",$p)*2; # unicode?
    $s="\0" x $l;
    $e=$MsiRecordGetString->Call($hdl,$num,$s,$l);
  }
  die $! if $e;

  $l=unpack("l",$p);
  return "((too big))" if ($l > length($s));

#  $l=index($s,"\0");
#  $l=length($s) if $l<0;
  return substr($s,0,$l);
}

sub _extract_fields
{
  my($hdl,@types)=@_;
  my(@a,$i,$l);

  $i=$MsiRecordGetFieldCount->Call($hdl);
  $i || die $!;
  @a=();
  $c=1;
  while ($c <= $i)
  {
    if (@types)
    {
      $l=shift @types;

      if ($l == $COLTYPE_INT)
      {
	push @a,_getI($hdl,$c);
      }
      elsif ($l == $COLTYPE_STRING)
      {
	push @a,_getS($hdl,$c);
      }
      else
      {
# STREAMS and other not processed here
	push @a,undef;
      }
    }
    else
    {
# autodetect-mode
      $s=_getI($hdl,$c);

      if (defined($s))
      {
	push @a,$s;
      }
      else
      {
	push @a,_getS($hdl,$c);
      }
    }

    $c++;
  }

  @a;
}

# vi rules :-)
# vim: sw=2 ai

