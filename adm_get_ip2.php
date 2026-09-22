<?php
function dns_get_record($host, $type='any'){ 
    $res=explode("\n",strstr(shell_exec('nslookup -type=any'.' '.escapeshellarg($host).' 4.2.2.4'),$host)); 
    $ret=array(); 
    $kn=array(); 
    $fin=array(); 
    $ft=array(); 
    foreach($res as $k=>$v){ 
        if(ereg($host,$v)){ 
            $kn[]=$k; 
        } 
        $exp=explode(' ',$v); 
        if($exp[1]=='TTL'){ 
            $ttl=$exp[3]; 
        } 
        if($exp[0]==$host.'    text'){ 
            $text=substr($res[$k+2],1); 
        } 
    } 
    foreach($kn as $k=>$v){ 
        $ex=explode(' ',$res[$v]); 
        $ret[$k]['target']=$ex[count($ex)-1]; 
        $ret[$k]['kn']=$v; 
    } 
    foreach($ret as $k=>$v){ 
        if($ret[$k]['target']==$host || !strpos($ret[$k]['target'],'.')){ 
            unset($ret[$k]); 
        }else{ 
            if($ttl){ 
                $ret[$k]['ttl']=$ttl; 
            } 
            $ret[$k]['class']='IN'; 
            $exp=explode(' ',$res[$ret[$k]['kn']]); 
            if($exp[0]==$host.'    internet'){ 
                $t='A'; 
                $ret[$k]['ip']=$ret[$k]['target']; 
                unset($ret[$k]['target']); 
            }else 
            if($exp[0]==$host.'    MX'){ 
                $t='MX'; 
                $ret[$k]['pri']=$exp[3]{0}; 
            }else 
            if($exp[0]==$host.'    nameserver'){ 
                $t='NS'; 
            } 
            $ret[$k]['type']=$t; 
            $ret[$k]['host']=$host; 
            unset($ret[$k]['kn']); 
            $fin[]=$ret[$k]; 
        } 
    } 
    if($text){ 
        $tno=count($fin); 
        $fin[$tno]['host']=$host; 
        $fin[$tno]['type']='TXT'; 
        $fin[$tno]['content']=$text; 
        $fin[$tno]['class']='IN'; 
        $fin[$tno]['ttl']=$ttl; 
    } 
    if($type!='any'){ 
        foreach($fin as $k=>$v){ 
            if($fin[$k]['type']!=strtoupper($type)){ 
                unset($fin[$k]); 
            }else{ 
                $ft[]=$fin[$k]; 
            } 
        } 
        $fin=$ft; 
    } 
    return $fin; 
} 


 print_r(dns_get_record('stillosg.ddns.net')); ?>