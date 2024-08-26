proc run_OS_tcl {N m mat_lib story_mat th_path SF dt mode_num has_damping zeta_mode zeta setting path gm_name NPTS g print_results} {
    
    proc myprint {print_results str} {
        if {$print_results == 1} {puts $str}
    }

    if {$mode_num >= 5} {set mode_num 5}

    wipe
    model basic -ndm 2 -ndf 3

    # node 
    node 1 0 0
    fix 1 1 1 1
    set all_node_tags [list 1]
    set story_nodes [list]
    for {set i 0} {$i < $N} {incr i} {
        node [expr $i + 2] 0 0 -mass [lindex $m $i] 0 0
        fix [expr $i + 2] 0 1 1
        lappend story_nodes [expr $i + 2]
        lappend all_node_tags [expr $i + 2]
    }
    set nodeTag [expr $i + 3]

    # material
    set f [open "$path/temp_NLMDOF_results/done.txt" w]
    puts $f 2
    close $f
    for {set i 0} {$i < [llength $mat_lib]} {incr i} {
        set mat [lrange [lindex $mat_lib $i] 0 end]
        uniaxialMaterial {*}$mat
    }
    set matTag [expr $i + 2]

    # element
    set element_tags [list]
    set all_element_tags [list]
    set current_ele_tag 1
    for {set i 0} {$i < $N} {incr i} {
        lappend element_tags [list]
        for {set j 0} {$j < [llength [lindex $story_mat $i]]} {incr j} {
            element zeroLength $current_ele_tag [expr $i + 1] [expr $i + 2] -mat [lindex [lindex $story_mat $i] $j] -dir 1 -doRayleigh 1
            set sub_list [lindex $element_tags $i]
            lappend sub_list $current_ele_tag
            lset element_tags $i $sub_list
            lappend all_element_tags $current_ele_tag
            incr current_ele_tag
        }
    }
    myprint $print_results "elements: $element_tags"
    myprint $print_results "all elements: $all_element_tags"

    # Eigen analysis
    if {$N > 5} {set solver -genBandArpack} {set solver -fullGenLapack}
    set lambda_ [eigen $solver $mode_num]
    set omg [list]
    for {set i 0} {$i < [llength $lambda_]} {incr i} {lappend omg double([expr [lindex $lambda_ $i] ** 0.5])}
    set T [list]
    set pi [expr 2 * asin(1)]
    for {set i 0} {$i < [llength $omg]} {incr i} {lappend T [expr 2 * $pi / [lindex $omg $i]]}
    myprint $print_results "Period:\n$T"
    set f [open [format "%s/temp_NLMDOF_results/Periods.txt" $path] w]
    foreach Ti $T {puts $f $Ti}
    close $f
  
    # ground motion
    timeSeries Path 1 -dt $dt -filePath $th_path -factor [expr $SF * $g]
    for {set i 0} {$i < $N} {incr i} {
        pattern Plain [expr $i + 1] 1 -fact [expr -[lindex $m $i]] {
            load [lindex $story_nodes $i] 1 0 0
        }
    }

    # Define the static node for obatining the base acceleration
    set large_m 0
    foreach m_i $m {
        if {$m_i > $large_m} {set large_m $m_i}
    }
    set large_m [expr $large_m * 1e6]
    node $nodeTag 0 0 0
    node [expr $nodeTag + 1] 0 0 0
    fix $nodeTag 1 1 1
    fix [expr $nodeTag + 1] 0 1 1
    mass [expr $nodeTag + 1] $large_m 0 0
    set static_node [expr $nodeTag + 1]
    uniaxialMaterial Elastic $matTag 0.0
    element zeroLength $current_ele_tag $nodeTag [expr $nodeTag + 1] -mat $matTag -dir 1 -doRayleigh 0
    pattern Plain [expr $i + 2] 1 -fact $large_m {
        load $static_node 1 0 0
    }
    incr nodeTag 2
    incr matTag
    incr current_ele_tag

    # Rayleigh damping
    if {$has_damping == 1} {
        set z1 double([lindex $zeta 0])
        set z2 double([lindex $zeta 1])
        if {$mode_num > 1} {
            set w1 [lindex $omg [expr [lindex $zeta_mode 0] - 1]]
            set w2 [lindex $omg [expr [lindex $zeta_mode 1] - 1]]
            set a [expr 2 * $w1 * $w2 / ($w2 ** 2 - $w1 ** 2) * ($w2 * $z1 - $w1 * $z2)]
            set b [expr 2 * $w1 * $w2 / ($w2 ** 2 - $w1 ** 2) * ($z2 / $w1 - $z1 / $w2)]
        } else {
            set w1 [lindex $omg 0]
            set a 0
            set b [expr 2 * $z1 / $w1]
        }
        region 1 -ele {*}$all_element_tags -rayleigh $a 0 $b 0
        region 1 -node {*}$all_node_tags -rayleigh $a 0 $b 0
        # rayleigh $a 0 $b 0
        myprint $print_results "Damping: a = $a, b = $b"
    } else {
        myprint $print_results "No damping"
    }

    # recorder
    file mkdir "$path/temp_NLMDOF_results"
    # 1 base node
    recorder Node -file [format "%s/temp_NLMDOF_results/%s_base_reaction.txt" $path $gm_name] -time -node 1 -dof 1 reaction
    recorder Node -file [format "%s/temp_NLMDOF_results/%s_base_acc.txt" $path $gm_name] -node $static_node -dof 1 accel
    recorder Node -file [format "%s/temp_NLMDOF_results/%s_base_vel.txt" $path $gm_name] -node $static_node -dof 1 vel
    recorder Node -file [format "%s/temp_NLMDOF_results/%s_base_disp.txt" $path $gm_name] -node $static_node -dof 1 disp
    # 2 floor nodes
    set floor_nodes [list]
    for {set i 0} {$i < $N} {incr i} {lappend floor_nodes [expr 2 + $i]}
    recorder Node -file [format "%s/temp_NLMDOF_results/%s_floor_acc.txt" $path $gm_name] -node {*}$floor_nodes -dof 1 accel
    recorder Node -file [format "%s/temp_NLMDOF_results/%s_floor_vel.txt" $path $gm_name] -node {*}$floor_nodes -dof 1 vel
    recorder Node -file [format "%s/temp_NLMDOF_results/%s_floor_disp.txt" $path $gm_name] -node {*}$floor_nodes -dof 1 disp
    # 3 material hysteretic curves
    recorder Element -file [format "%s/temp_NLMDOF_results/%s_material.txt" $path $gm_name] -ele {*}$all_element_tags material 1 stressStrain
    # 4 modal results
    for {set i 1} {$i < [expr $mode_num + 1]} {incr i} {
        recorder Node -file [format "%s/temp_NLMDOF_results/mode_%d.txt" $path $i] -node {*}$floor_nodes -dof 1 "eigen $i"
    }

    # Time history analysis
    if {[lindex $setting 6] != ""} {
        constraints [lindex $setting 0] [lindex $setting 6] [lindex $setting 7]
    } else {
        constraints [lindex $setting 0]
    }
    numberer [lindex $setting 1]
    system [lindex $setting 2]
    test [lindex $setting 3] [lindex $setting 8] [lindex $setting 9]
    algorithm [lindex $setting 4]
    integrator [lindex $setting 5] [lindex $setting 10] [lindex $setting 11]
    analysis Transient
    
    set current_time 0
    set duration [expr $dt * ($NPTS - 1)]
    set init_dt $dt
    set factor 1.0
    set max_factor [lindex $setting 12]
    set min_factor [lindex $setting 13]
    set dt_ratio [lindex $setting 14]
    set done 0

    while 1 {
        if {$current_time >= $duration} {
            set done 1
            break;  # analysis finished
        }
        set dt [expr $init_dt * $factor * $dt_ratio]
        if {$current_time + $dt > $duration} {
            set dt [expr $duration - $current_time]
        }
        set ok [analyze 1 $dt]
        if {$ok == 0} {
            # current step finished
            set current_time [expr $current_time + $dt]
            set old_factor $factor
            set factor [expr $factor * 2]
            set factor [expr min($factor, $max_factor)]
            set dt [expr $init_dt * $factor]
            if {$factor != $old_factor} {
                myprint $print_results "--- Enlarge factor to $factor ---"
            }
        } else {
            # current step did not converge
            set factor [expr double($factor) / 4]
            if {$factor < $min_factor} {
                myprint $print_results "--- factor is less than the minimum allowed ($factor < $min_factor). ---"
                myprint $print_results "--- current time: $current_time. total time: $duration. ---"
                myprint $print_results "--- The analysis did not converge. ---"
                set done 0
                break;  # analysis failed
            } else {
                # reduce factor
                set dt [expr $init_dt * $factor]
                myprint $print_results "--- Current step did not converge, reduce factor to $factor ---"
            }
        }
    }
    wipeAnalysis
    wipe
    if {$done == 1} {
        myprint $print_results "------ Finished ------"
    } else {
        myprint $print_results "------ Not converge ------"
    }
    set f [open "$path/temp_NLMDOF_results/done.txt" w]
    if {$done == 1} {puts $f 1} else {puts $f 0}
    close $f
}


if {$argc == 0} {
    set N 3
    set m [list 2 1 1]
    set mat_lib [list [list Steel01 1 3000 1500 0.02] [list Steel01 2 2000 1000 0.02]]
    set story_mat [list [list 1] [list 2] [list 2]]
    set th_path "test/ChiChi.txt";  # replace "xxx" with groumd motion path1
    set SF 1
    set dt 0.01
    set mode_num 3
    set has_damping 1
    set zeta_mode [list 1 2]
    set zeta [list 0.05 0.05]
    set setting [list Transformation Plain BandGeneral NormUnbalance Newton Newmark "" "" 1e-5 60 0.5 0.25 1 1e-6 1]
    set path "temp"
    set gm_name "ChiChi"
    set NPTS 5279
    set g 9810.0
    set print_results 0
    run_OS_tcl $N $m $mat_lib $story_mat $th_path $SF $dt $mode_num $has_damping $zeta_mode $zeta $setting $path $gm_name $NPTS $g $print_results
}
