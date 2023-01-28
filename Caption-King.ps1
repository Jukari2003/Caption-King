################################################################################
#                               Caption King                                   #
#                     Written By: MSgt Anthony Brechtel                        #
#                                                                              #
################################################################################
######Load Assemblies###########################################################
clear-host
Add-Type -AssemblyName 'System.Windows.Forms'
Add-Type -AssemblyName 'System.Drawing'
Add-Type -AssemblyName 'PresentationFramework'
[System.Windows.Forms.Application]::EnableVisualStyles();
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")

################################################################################
######Global Variables##########################################################
$scriptpath = $MyInvocation.MyCommand.Path
$dir = Split-Path $scriptpath
Set-Location $dir
$script:debug = 0;
$script:settings = @{};
$script:version = "1.3"
$script:program_title = "Caption King"
$script:last_picture = 0;
$script:current_picture = 1;
$script:caption_path = ""
$script:img = ""
$script:started = 0;
$script:click_active = 0;

$script:file_list = ""
$script:file_count = 0;
$script:good_up_to = 0;
$script:wait = 0;
$script:wait_mode = "";
$script:canvas_x = 0;
$script:canvas_y = 0;
$script:scale  = 0;
$script:Form                    = New-Object System.Windows.Forms.Form
$title1                         = New-Object System.Windows.Forms.Label 
$title2                         = New-Object System.Windows.Forms.Label
$directory_label                = New-Object System.Windows.Forms.Label
$target_box                     = New-Object System.Windows.Forms.TextBox
$script:image_location_trackbar = New-Object System.Windows.Forms.TrackBar
$image_number_label             = New-Object System.Windows.Forms.Label
$browse_button                  = New-Object System.Windows.Forms.Button
$script:picture_box             = New-Object system.Windows.Forms.Panel
$script:previous_button1        = New-Object System.Windows.Forms.Button
$script:previous_button2        = New-Object System.Windows.Forms.Button
$script:next_button1            = New-Object System.Windows.Forms.Button
$script:next_button2            = New-Object System.Windows.Forms.Button
$script:caption_box             = New-Object system.Windows.Forms.TextBox
$script:caption_box_combo       = New-Object System.Windows.Forms.ComboBox
$script:revert_button           = New-Object System.Windows.Forms.Button
$script:working_dir_button      = New-Object System.Windows.Forms.Button
$script:wrong_dim_button        = New-Object System.Windows.Forms.Button
$script:no_captions_button      = New-Object System.Windows.Forms.Button
$script:source_dir_button	    = New-Object System.Windows.Forms.Button
$script:exempt_button           = New-Object System.Windows.Forms.Button
$image_name_label1              = New-Object System.Windows.Forms.Label
$rotate_right_button            = New-Object System.Windows.Forms.Button
$rotate_left_button             = New-Object System.Windows.Forms.Button
$image_name_label2              = New-Object System.Windows.Forms.Label
$dimensions_label1              = New-Object System.Windows.Forms.Label
$dimensions_label2              = New-Object System.Windows.Forms.Label
$script:dim_height_input        = New-Object System.Windows.Forms.TextBox
$script:dim_x_label             = New-Object System.Windows.Forms.Label
$script:dim_width_input         = New-Object System.Windows.Forms.TextBox
$script:target_dim_label        = New-Object System.Windows.Forms.Label
$wait_label                     = New-Object System.Windows.Forms.Label
$no_images_label                = New-Object System.Windows.Forms.Label

$script:caption_list            = New-Object system.collections.hashtable
$script:rack_and_stack = @{};

$script:output_dir = "";                        #System Required Directory
$script:user_output_dir = "$dir\Working"        #User Defined Working Location


$script:crosshair_size_x = $script:settings['TARGET_WIDTH']
$script:crosshair_size_y = $script:settings['TARGET_HEIGHT']
$script:crosshair_x1 = 0
$script:crosshair_x2 = 0
$script:crosshair_y1 = 0
$script:crosshair_y2 = 0
$script:target_lock = 0;

$script:image_height = 0        #Actual Size
$script:image_width = 0         #Actual Size
$script:image_height_s = 50     #Scaled to Canvas
$script:image_width_s = 50      #Scaled to Canvas

$script:image_location_point_corner1x = 0 #Image Canvas Boundries
$script:image_location_point_corner1y = 0 #Image Canvas Boundries
$script:image_location_point_corner2x = 0 #Image Canvas Boundries
$script:image_location_point_corner2y = 0 #Image Canvas Boundries


$script:form_width = 1300
$script:form_height = 1200
$script:lock = 0;
##Idle Timer
if(Test-Path variable:Script:Timer){$Script:Timer.Dispose();}
$Script:Timer = New-Object System.Windows.Forms.Timer
$Script:Timer.Interval = 1000
$Script:CountDown = 1
################################################################################
######Main######################################################################
function main
{
    
    ##################################################################################
    ###########Main Form
    $script:Form.Font = "Copperplate Gothic,8.1"
    $script:Form.ForeColor = "Black"
    $script:Form.BackColor = "#434343"
    $script:Form.Text = "  $script:program_title"
    $script:Form.Width = $script:form_width
    $script:Form.Height = $script:form_height
    $script:Form.MinimumSize  = New-Object Drawing.Size(1280,800)
    $y_pos = 0
   

    ##################################################################################
    ###########Directory Label
    $y_pos = $y_pos + 15
    $target_box.width = 385
    $browse_button.Width=70
    $directory_label.Size = "150, 23"
    $directory_label.Location = New-Object System.Drawing.Point((($script:Form.width / 2) - (($target_box.width + $browse_button.Width + $directory_label.width)/ 2)),$y_pos)
    $directory_label.ForeColor = "White" 
    $directory_label.Text = "Source Directory:"
    $directory_label.TextAlign = "Middleleft"
    $directory_label.Font = New-Object System.Drawing.Font("Copperplate Gothic",10,[System.Drawing.FontStyle]::Bold)
    

    ##################################################################################
    ###########Target Box   
    $target_box.Location = New-Object System.Drawing.Point(($directory_label.location.x + $directory_label.width + 3),($y_pos))
    $target_box.font = New-Object System.Drawing.Font("Arial",11,[System.Drawing.FontStyle]::Regular)
    $target_box.Height = 40
    $target_box.Text = $script:settings['TARGET_DIRECTORY']
    if($started -eq 0)
    {
        $target_box.Add_Click({
            if($this.Text -eq "Browse or Enter a file path")
            {
                $this.Text = ""
            }
        })
        $target_box.Add_lostFocus({
            if($this.text -eq "Browse or Enter a file path")
            {

            }
            elseif(($this.text -eq "") -or ($this.text -eq $null))
            {
                $this.text = "Browse or Enter a file path"
            }
            elseif(!(Test-Path -literalpath $this.text))
            {
                $this.text = "Browse or Enter a file path"
            }
            elseif(!((Get-Item -literalpath $this.text) -is [System.IO.DirectoryInfo]))   
            {
                $this.text = "Browse or Enter a file path"
            }
        
        })
        $target_box.Add_TextChanged({
            $this.text = $this.text -replace "^`"|`"$"
            $script:settings['TARGET_DIRECTORY'] = $this.text
            if(($script:settings['TARGET_DIRECTORY'] -ne $null) -and ($script:settings['TARGET_DIRECTORY'] -ne ""))
            {
                if(Test-Path -literalpath $target_box.text)
                {
                    $script:settings['TARGET_DIRECTORY'] = $this.text
                    update_settings
                    $script:last_picture = 0;
                    $script:current_picture = 1;
                    $image_location_trackbar.Value = 1
                    $script:output_dir = ""
                }
            }
            else
            {
                $script:Form.Controls.Remove($scan_target_button)
            }
        })
    }#First Run
    

    ##################################################################################
    ###########Browse Button   
    $browse_button.Location= New-Object System.Drawing.Size(($target_box.location.x + $target_box.width + 3),($y_pos - 3))
    $browse_button.BackColor = "#606060"
    $browse_button.ForeColor = "White"
    $browse_button.Height=30
    $browse_button.Text='Browse'
    $browse_button.Font = [Drawing.Font]::New("Times New Roman", 9)
    if($started -eq 0)
    {
        $browse_button.Add_Click(
        {    
		    prompt_for_folder
            if(($script:settings['TARGET_DIRECTORY'] -ne $Null) -and ($script:settings['TARGET_DIRECTORY'] -ne "") -and ((Test-Path -literalpath $script:settings['TARGET_DIRECTORY']) -eq $True))
            {
                $target_box.Text= $script:settings['TARGET_DIRECTORY']
            }
        })
    }
    

    ##################################################################################
    ###########Image Location Trackbar
    $y_pos = $y_pos + 45
    $image_location_trackbar.Width = $script:Form.Width - 20
    $image_location_trackbar.Location = New-Object System.Drawing.Size((($script:Form.width / 2) - ($image_location_trackbar.Width / 2) - 8),($y_pos))  
    $image_location_trackbar.Orientation = "Horizontal"
    $image_location_trackbar.Height = 10
    $image_location_trackbar.LargeChange = 20
    $image_location_trackbar.SmallChange = 1
    $image_location_trackbar.TickFrequency = 200
    $image_location_trackbar.TickStyle = "TopLeft"
    $image_location_trackbar.SetRange(1, 3000)

    if($started -eq 0)
    {
        $image_location_trackbar.Value = 1
        $image_location_trackbar.add_ValueChanged({
            $Script:Timer.Interval = 350
            $image_number_label.Text = [string]$this.value + " of $script:file_count"
            $script:wait_mode = "Trackbar";
            $script:wait = $Script:CountDown + 1   
        })
    }
    

    ##################################################################################
    ###########Image Number Label
    $image_number_label.Font       = New-Object System.Drawing.Font("Copperplate Gothic",10,[System.Drawing.FontStyle]::Bold)
    $image_number_label.Text       = ""
    $image_number_label.TextAlign  = "MiddleCenter"
    $image_number_label.ForeColor  = "White"
    $image_number_label.Width      = $script:Form.Width
    $image_number_label.Height     = 20
    $image_number_label.Location   = New-Object System.Drawing.Size((($script:Form.Width / 2) / ($image_number_label.Width / 2)),($y_pos - 20))


    ##################################################################################
    ###########Picture Box
    $y_pos = $y_pos + 50
    $picture_box.Size = New-Object System.Drawing.Size(($script:Form.width - 130),($script:Form.height - 350))
    $picture_box.Location = New-Object System.Drawing.Size((($script:Form.width / 2) - ($picture_box.width / 2) - 8),($y_pos + 1))  
    $picture_box.backcolor = "Black"    
    if($started -eq 0)
    {
        $script:picture_box.Add_MouseDown({
            if($_.Button -eq [System.Windows.Forms.MouseButtons]::Left ) 
            {
                if($script:click_active -eq 0)
                {
                    $script:click_active = 1;
                    $Script:Timer.Interval = 80
                }
                else
                {
                    $script:click_active = 0;
                    $Script:Timer.Interval = 500
                    process_crosshair
                }
            }
            if($_.Button -eq [System.Windows.Forms.MouseButtons]::Right ) 
            {
                $script:target_lock = 1
                $script:click_active = 0;
                $Script:Timer.Interval = 500
                [string]$image = [io.path]::GetFileNameWithoutExtension($image_path) + [System.IO.Path]::GetExtension($image_path)
                load_picture $image #Reset Image
            }

        })
        $script:picture_box.add_MouseWheel({
            $m = [System.Windows.Forms.MouseEventArgs]$_
            $script:zoom = 1
            $script:click_active = 1
            $increment = 10
            $Script:Timer.Interval = 50
            #######################################################
            ####Zoom Up
            if($m.delta -ge 0)
            {
                determine_crosshair_size "increase"
            }
            #######################################################
            ####Zoom Down
            else
            {
                determine_crosshair_size "decrease"
            }
            $script:zoom = 0
        })
    }#First Run


    ##################################################################################
    ###########Previous Image Button 1
    $previous_button1.Width= 55  
    $previous_button1.Location= New-Object System.Drawing.Size(($picture_box.Location.x - $previous_button1.Width),($picture_box.Location.y + 55))
    $previous_button1.BackColor = "#606060"
    $previous_button1.ForeColor = "White"
    $previous_button1.Height= (($picture_box.height - 55) / 2)
    $previous_button1.Text="<"
    $previous_button1.Font = [Drawing.Font]::New("Times New Roman", 35)
    if($started -eq 0)
    {
        $previous_button1.Add_Click(
        {   
            $script:click_active = 0;
            $Script:Timer.Interval = 500
            if($caption_box.text -ne $caption_box.AccessibleDescription)
            {
                $caption_box.AccessibleDescription = $caption_box.text
                Set-Content $script:caption_path_working $caption_box.text
            }
            $script:current_picture--;
            $image_location_trackbar.Value = $script:current_picture         
        })
    }


    ##################################################################################
    ###########Previous Image Button 1
    $previous_button2.Width= 55  
    $previous_button2.Location= New-Object System.Drawing.Size(($picture_box.Location.x - $previous_button1.Width),($previous_button1.Location.y + $previous_button1.Height))
    $previous_button2.BackColor = "#606060"
    $previous_button2.ForeColor = "White"
    $previous_button2.Height= (($picture_box.height - 55) / 2)
    $previous_button2.Text="<`n⚖"
    $previous_button2.Font = [Drawing.Font]::New("Times New Roman", 35)
    if($started -eq 0)
    {
        $previous_button2.Add_Click(
        {   
            $script:click_active = 0;
            $Script:Timer.Interval = 500
            if($caption_box.text -ne $caption_box.AccessibleDescription)
            {
                $caption_box.AccessibleDescription = $caption_box.text
                Set-Content $script:caption_path_working $caption_box.text
            }
            resize_image
            $script:wait_mode = "Prev"         
        })
    }
    ##################################################################################
    ###########Rotate Right   
    $rotate_left_button.Location= New-Object System.Drawing.Size(($picture_box.Location.x - $previous_button1.Width),$picture_box.Location.y)
    $rotate_left_button.BackColor = "#606060"
    $rotate_left_button.ForeColor = "White"
    $rotate_left_button.Width= 55
    $rotate_left_button.Height= 55
    $rotate_left_button.Text='⟲'
    $rotate_left_button.TextAlign = "MiddleCenter"
    $rotate_left_button.Font = [Drawing.Font]::New("Times New Roman", 35)
    if($started -eq 0)
    {
        $rotate_left_button.Add_Click(
        {
            $script:click_active = 0;
            $Script:Timer.Interval = 50  
            rotate_image "left"
            #$Script:Timer.Interval = 1000    
        })
    }


    ##################################################################################
    ###########Rotate Right   
    $rotate_right_button.Location= New-Object System.Drawing.Size(($picture_box.Location.x + $picture_box.width),$picture_box.Location.y)
    $rotate_right_button.BackColor = "#606060"
    $rotate_right_button.ForeColor = "White"
    $rotate_right_button.Width= 55
    $rotate_right_button.Height= 55
    $rotate_right_button.Text='⟳'
    $rotate_right_button.TextAlign = "MiddleCenter"
    $rotate_right_button.Font = [Drawing.Font]::New("Times New Roman", 35)
    if($started -eq 0)
    {
        
        $rotate_right_button.Add_Click(
        {
            $script:click_active = 0;
            $Script:Timer.Interval = 50
            rotate_image "right"
            #$Script:Timer.Interval = 1000      
        })
    }
    ##################################################################################
    ###########next Image button1   
    $next_button1.Location= New-Object System.Drawing.Size(($picture_box.Location.x + $picture_box.width),($picture_box.Location.y + 55))
    $next_button1.BackColor = "#606060"
    $next_button1.ForeColor = "White"
    $next_button1.Width= 55
    $next_button1.Height= (($picture_box.height - 55) / 2)
    $next_button1.Text=">"
    $next_button1.Font = [Drawing.Font]::New("Times New Roman", 35)
    if($started -eq 0)
    {
        $next_button1.Add_Click(
        {
            $script:click_active = 0;
            $Script:Timer.Interval = 500   
            if($caption_box.text -ne $caption_box.AccessibleDescription)
            {
                $caption_box.AccessibleDescription = $caption_box.text
                Set-Content $script:caption_path_working $caption_box.text
            }
            $script:current_picture++;
            $image_location_trackbar.Value = $script:current_picture
        })
    }

    ##################################################################################
    ###########next Image button2   
    $next_button2.Location= New-Object System.Drawing.Size(($picture_box.Location.x + $picture_box.width),($next_button1.Location.y + $next_button1.Height))
    $next_button2.BackColor = "#606060"
    $next_button2.ForeColor = "White"
    $next_button2.Width= 55
    $next_button2.Height= (($picture_box.height - 55) / 2)
    $next_button2.Text=">`n⚖"
    $next_button2.Font = [Drawing.Font]::New("Times New Roman", 35)
    if($started -eq 0)
    {
        $next_button2.Add_Click(
        {   
            $script:click_active = 0;
            $Script:Timer.Interval = 500
            if($caption_box.text -ne $caption_box.AccessibleDescription)
            {
                $caption_box.AccessibleDescription = $caption_box.text
                Set-Content $script:caption_path_working $caption_box.text
            }
            resize_image
            $script:wait_mode = "Next"
        })
    }


    ##################################################################################
    ###########Caption Box
    $y_pos = $y_pos + $picture_box.height + 5
    $caption_box.Font                           = New-Object System.Drawing.Font("Arial",14,[System.Drawing.FontStyle]::Regular)
    $caption_box.Size                           = New-Object System.Drawing.Size(($previous_button1.width + $next_button1.width + $picture_box.Width),60)
    $caption_box.Location                       = New-Object System.Drawing.Size($previous_button1.Location.x,$y_pos)    
    $caption_box.WordWrap                       = $true
    $caption_box.Multiline                      = $True
    $caption_box.BackColor                      = "white"
    $caption_box.ScrollBars                     = "vertical"

    ##################################################################################
    ###########Caption Drop Down
    $y_pos = $y_pos + $caption_box.height + 2;
    $script:caption_box_combo.width = ($caption_box.width)
    $script:caption_box_combo.autosize = $false
    $script:caption_box_combo.Anchor = 'top,right'
    $script:caption_box_combo.Font                           = New-Object System.Drawing.Font("Arial",14,[System.Drawing.FontStyle]::Regular)
    $script:caption_box_combo.Location = New-Object System.Drawing.Point($caption_box.Location.x,$y_pos)
    $script:caption_box_combo.DropDownStyle = "DropDownList"
    $first = "Quick Add Drop Down";
    if($started -eq 0)
    {
        $script:caption_box_combo.Add_SelectedValueChanged(
        {
            if($script:lock -ne 1)
            {
                $script:caption_box.text = $this.SelectedItem
            }
        })
        $script:caption_box_combo.Add_MouseDown(
        {
            if($_.Button -eq [System.Windows.Forms.MouseButtons]::Right ) 
            {
                if($script:lock -ne 1)
                {
                    $script:caption_box.text = $this.SelectedItem
                }
            }
        })
    }
    ##################################################################################
    ###########Revert to Original Button
    $y_pos = $y_pos + 35
    
    $revert_button.BackColor = "#606060"
    $revert_button.ForeColor = "White"
    $revert_button.Width= 200
    $revert_button.Height= 30
    $revert_button.Text='Revert to Original'
    $revert_button.Location= New-Object System.Drawing.Size((($form.Width / 2)), $y_pos)
    $revert_button.TextAlign  = "MiddleCenter"
    $revert_button.Font = [Drawing.Font]::New("Times New Roman", 9)
    if($started -eq 0)
    {
        $revert_button.Add_Click({   
            $script:bitmap.Dispose()
            [System.GC]::GetTotalMemory(‘forcefullcollection’) | out-null
            Remove-Item -LiteralPath $script:image_path_working
            $script:last_picture = 0;
        })
    }
    ##################################################################################
    ###########Working Directory   
    $working_dir_button.BackColor = "#606060"
    $working_dir_button.ForeColor = "White"
    $working_dir_button.Width= 200
    $working_dir_button.Height= 30
    $working_dir_button.Text='Working Directory'
    $working_dir_button.Location= New-Object System.Drawing.Size(($revert_button.Location.x - ($working_dir_button.width + 5)), $y_pos)
    $working_dir_button.TextAlign  = "MiddleCenter"
    $working_dir_button.Font = [Drawing.Font]::New("Times New Roman", 9)
    if($started -eq 0)
    {
        $working_dir_button.Add_Click({   
            explorer.exe $script:output_dir
        })
    }

    ##################################################################################
    ###########Source Directory   
    $source_dir_button.BackColor = "#606060"
    $source_dir_button.ForeColor = "White"
    $source_dir_button.Width= 200
    $source_dir_button.Height= 30
    $source_dir_button.Text='Source Directory'
    $source_dir_button.Location= New-Object System.Drawing.Size(($working_dir_button.Location.x - ($working_dir_button.width + 5)), $y_pos)
    $source_dir_button.TextAlign  = "MiddleCenter"
    $source_dir_button.Font = [Drawing.Font]::New("Times New Roman", 9)
    if($started -eq 0)
    {
        $source_dir_button.Add_Click({   
            explorer.exe $script:settings['TARGET_DIRECTORY']
        })
    }

    ##################################################################################
    ###########Exempt Button   
    $exempt_button.BackColor = "#606060"
    $exempt_button.ForeColor = "White"
    $exempt_button.Width= 200
    $exempt_button.Height= 30
    $exempt_button.Text='Ignore Image'
    $exempt_button.Location= New-Object System.Drawing.Size(($source_dir_button.Location.x - ($source_dir_button.width + 5)), $y_pos)
    $exempt_button.TextAlign  = "MiddleCenter"
    $exempt_button.Font = [Drawing.Font]::New("Times New Roman", 9)
    if($started -eq 0)
    {
        $exempt_button.Add_Click({   
            Add-Content -LiteralPath ($script:output_dir + "\Ignored.csv") $script:image_file.FullName
            if(Test-Path -LiteralPath $script:image_path_working)
            {
                if(Test-Path variable:bitmap){$script:bitmap.Dispose()}
                [System.GC]::GetTotalMemory(‘forcefullcollection’) | out-null
                Remove-Item -literalPath $script:image_path_working
            }
            if(Test-Path -LiteralPath $script:caption_path_working)
            {
                Remove-Item -Literalpath $script:caption_path_working
            }
            $script:current_picture++
        })
    }

    ##################################################################################
    ###########Only Wrong Dimensions 
    $wrong_dim_button.BackColor = "#606060"
    $wrong_dim_button.ForeColor = "White"
    $wrong_dim_button.Width= 200
    $wrong_dim_button.Height= 30
    $wrong_dim_button.Text='Find Wrong Dimensions'
    $wrong_dim_button.Location= New-Object System.Drawing.Size(($revert_button.Location.x + ($revert_button.width + 5)), $y_pos)
    $wrong_dim_button.TextAlign  = "MiddleCenter"
    $wrong_dim_button.Font = [Drawing.Font]::New("Times New Roman", 9)
    if($started -eq 0)
    {
        if($script:settings['FIND_DIMENSIONS'] -eq "On")
        {
            $wrong_dim_button.ForeColor = "#39ff14"
        }

        $wrong_dim_button.Add_Click({   
            if($script:settings['FIND_DIMENSIONS'] -eq "Off")
            {
                $script:settings['FIND_DIMENSIONS'] = "On";
                update_settings
                $wrong_dim_button.ForeColor = "#39ff14"
                $script:last_picture = 0;
                $script:current_picture = 1
            }
            else
            {
                $script:good_up_to = 0
                $script:settings['FIND_DIMENSIONS'] = "Off";
                update_settings
                $wrong_dim_button.ForeColor = "White"
            }
        })
    }
    ##################################################################################
    ###########Only No Captions
    $no_captions_button.BackColor = "#606060"
    $no_captions_button.ForeColor = "White"
    $no_captions_button.Width= 200
    $no_captions_button.Height= 30
    $no_captions_button.Text='Find Missing Captions'
    $no_captions_button.Location= New-Object System.Drawing.Size(($wrong_dim_button.Location.x + ($wrong_dim_button.width + 5)), $y_pos)
    $no_captions_button.TextAlign  = "MiddleCenter"
    $no_captions_button.Font = [Drawing.Font]::New("Times New Roman", 9)
    if($started -eq 0)
    {
        if($script:settings['FIND_CAPTIONS'] -eq "On")
        {
            $no_captions_button.ForeColor = "#39ff14"
        }

        $no_captions_button.Add_Click({   
            if($script:settings['FIND_CAPTIONS'] -eq "Off")
            {
                $script:settings['FIND_CAPTIONS'] = "On";
                update_settings
                $no_captions_button.ForeColor = "#39ff14"
                $script:last_picture = 0;
                $script:current_picture = 1

            }
            else
            {
                $script:good_up_to = 0
                $script:settings['FIND_CAPTIONS'] = "Off";
                update_settings
                $no_captions_button.ForeColor = "White"
            }
        })
    }


    ##################################################################################
    ###########Dimensions1
    $y_pos = $y_pos + 30
    $dimensions_label1.Font       = New-Object System.Drawing.Font("Copperplate Gothic",10,[System.Drawing.FontStyle]::Bold)
    $dimensions_label1.Text       = "Dimensions:"
    $dimensions_label1.TextAlign  = "MiddleRight"
    $dimensions_label1.ForeColor  = "White"
    $dimensions_label1.Width      = 150
    $dimensions_label1.Height     = 40
    $dimensions_label1.Location   = New-Object System.Drawing.Size((($script:Form.Width / 2) - (500 / 2)),$y_pos)


    ##################################################################################
    ###########Dimensions2
    $dimensions_label2.Font       = New-Object System.Drawing.Font("Copperplate Gothic",9.5,[System.Drawing.FontStyle]::Regular)
    $dimensions_label2.Text       = ""
    $dimensions_label2.TextAlign  = "MiddleLeft"
    $dimensions_label2.ForeColor  = "darkgray"
    $dimensions_label2.Width      = 100
    $dimensions_label2.Height     = 40
    $dimensions_label2.Location   = New-Object System.Drawing.Size(($dimensions_label1.location.x + $dimensions_label1.width),$y_pos)


    ##################################################################################
    ###########Image Name1
    $image_name_label1.Font       = New-Object System.Drawing.Font("Copperplate Gothic",10,[System.Drawing.FontStyle]::Bold)
    $image_name_label1.Text       = "Image Name:"
    $image_name_label1.TextAlign  = "MiddleRight"
    $image_name_label1.ForeColor  = "white"
    $image_name_label1.Width      = 150
    $image_name_label1.Height     = 40
    $image_name_label1.Location  = New-Object System.Drawing.Size(($dimensions_label2.location.x + $dimensions_label2.width),$y_pos)


    ##################################################################################
    ###########Image Name2
    $image_name_label2.Font       = New-Object System.Drawing.Font("Copperplate Gothic",9.5,[System.Drawing.FontStyle]::Regular)
    $image_name_label2.Text       = ""
    $image_name_label2.TextAlign  = "MiddleLeft"
    $image_name_label2.ForeColor  = "darkgray"
    $image_name_label2.Width      = 400
    $image_name_label2.Height     = 40
    $image_name_label2.Location   = New-Object System.Drawing.Size(($image_name_label1.location.x + $image_name_label1.width),$y_pos)


    ##################################################################################
    ###########No Images 
    $no_images_label.Font       = New-Object System.Drawing.Font("Copperplate Gothic",50,[System.Drawing.FontStyle]::Regular)
    $no_images_label.Text       = "No Images"
    $no_images_label.TextAlign  = "MiddleCenter"
    $no_images_label.ForeColor  = "Red"
    $no_images_label.backcolor = "black"
    $no_images_label.Width      = $picture_box.width
    $no_images_label.Height     = $picture_box.height
    $no_images_label.Location   = New-Object System.Drawing.Size($picture_box.Location.x,$picture_box.Location.y)
    $no_images_label.hide();

    ##################################################################################
    ###########No Images 
    $wait_label.Font       = New-Object System.Drawing.Font("Copperplate Gothic",25,[System.Drawing.FontStyle]::Regular)
    $wait_label.Text       = "3"
    $wait_label.TextAlign  = "TopCenter"
    $wait_label.ForeColor  = "Red"
    $wait_label.Width      = 60
    $wait_label.Height     = 60
    $wait_label.Location   = New-Object System.Drawing.Size(($form.width - ($wait_label.Width + 40)),0)
    $wait_label.hide();


    ##################################################################################
    ###########Set Height Dimension Input
    $dim_height_input.width = 40
    $dim_height_input.Location = New-Object System.Drawing.Point(($dimensions_label1.Location.x - ($dim_height_input.width + 5)),($y_pos + 8))
    $dim_height_input.font = New-Object System.Drawing.Font("Arial",11,[System.Drawing.FontStyle]::Regular)
    $dim_height_input.Height = 40
    $dim_height_input.Text = $script:settings['TARGET_HEIGHT']
    if($started -eq 0)
    {
        $dim_height_input.Add_lostFocus({
            if(($this.text -match "^\d+$") -and ($this.text.length -le 4))
            {
                $script:settings['TARGET_HEIGHT'] = $this.text
                update_settings
                determine_crosshair_size "image"
            }
            else
            {
                $this.text = $script:settings['TARGET_HEIGHT']
            }
         })
        $dim_height_input.Add_TextChanged({
            if(($this.text -match "^\d+$") -and ($this.text.length -le 4))
            {
                $script:settings['TARGET_HEIGHT'] = $this.text
                update_settings
                determine_crosshair_size "image"
            }
            else
            {
                $this.text = $script:settings['TARGET_HEIGHT']
            }
        })
    }#First Run
    
    
    ##################################################################################
    ###########X Label 
    $dim_x_label.Font       = New-Object System.Drawing.Font("Copperplate Gothic",9.5,[System.Drawing.FontStyle]::Regular)
    $dim_x_label.Text       = "x"
    $dim_x_label.TextAlign  = "MiddleCenter"
    $dim_x_label.ForeColor  = "white"
    $dim_x_label.Width      = 20
    $dim_x_label.Height     = 20
    $dim_x_label.Location   = New-Object System.Drawing.Size(($dim_height_input.Location.x - ($dim_x_label.width)),($y_pos + 11))

    ##################################################################################
    ###########Set Width Dimension Input
    $dim_width_input.width = 40
    $dim_width_input.Location = New-Object System.Drawing.Point(($dim_x_label.Location.x - ($dim_width_input.width)),($y_pos + 8))
    $dim_width_input.font = New-Object System.Drawing.Font("Arial",11,[System.Drawing.FontStyle]::Regular)
    $dim_width_input.width = 40
    $dim_width_input.Text = $script:settings['TARGET_WIDTH']
    if($started -eq 0)
    {
        $dim_width_input.Add_lostFocus({
            if(($this.text -match "^\d+$") -and ($this.text.length -le 4))
            {
                $script:settings['TARGET_WIDTH'] = $this.text
                update_settings
                determine_crosshair_size "image"
            }
            else
            {
                $this.text = $script:settings['TARGET_WIDTH']
                update_settings
            }
         })
        $dim_width_input.Add_TextChanged({
            if(($this.text -match "^\d+$") -and ($this.text.length -le 4))
            {
                $script:settings['TARGET_WIDTH'] = $this.text
                determine_crosshair_size "image"
            }
            else
            {
                $this.text = $script:settings['TARGET_WIDTH']
            }
        })
    }#First Run

    ##################################################################################
    ###########Target Dim Label
    $target_dim_label.Font       = New-Object System.Drawing.Font("Copperplate Gothic",10,[System.Drawing.FontStyle]::Bold)
    $target_dim_label.Text       = "Target Dimensions:"
    $target_dim_label.TextAlign  = "MiddleRight"
    $target_dim_label.ForeColor  = "White"
    $target_dim_label.Width      = 200
    $target_dim_label.Height     = 40
    $target_dim_label.Location   = New-Object System.Drawing.Size(($dim_width_input.Location.x - ($target_dim_label.width + 5)),($y_pos))


    ##################################################################################
    ###########First Run Load Items
    $script:formGraphics            = $script:picture_box.createGraphics()
    if($script:started -eq 0)
    {
        $script:started = 1;
        $Script:Timer.Start()
        $Script:Timer.Add_Tick({Idle_Timer})
        $script:Form.Controls.Add($directory_label)
        $script:Form.Controls.Add($target_box)
        $script:Form.Controls.Add($browse_button)
        $script:Form.Controls.Add($previous_button1)
        $script:Form.Controls.Add($previous_button2)
        $script:Form.Controls.Add($next_button1)
        $script:Form.Controls.Add($next_button2)
        $script:Form.Controls.Add($rotate_right_button)
        $script:Form.Controls.Add($rotate_left_button)
        $script:Form.Controls.Add($caption_box)
        $script:Form.Controls.Add($script:caption_box_combo)
        $script:Form.Controls.Add($script:revert_button)
        $script:Form.Controls.Add($script:working_dir_button)
        $script:Form.Controls.Add($script:source_dir_button)
        $script:Form.Controls.Add($script:exempt_button) 
        $script:Form.Controls.Add($script:wrong_dim_button)
        $script:Form.Controls.Add($script:no_captions_button)
        $script:Form.Controls.add($dimensions_label1)
        $script:Form.Controls.add($dimensions_label2)
        $script:Form.Controls.add($image_name_label1)
        $script:Form.Controls.add($image_name_label2)
        $script:Form.Controls.add($script:dim_height_input) 
        $script:Form.Controls.add($dim_x_label)
        $script:Form.Controls.add($script:dim_width_input)
        $script:Form.Controls.add($target_dim_label)
        $script:Form.controls.add($script:picture_box)
        $script:Form.Controls.add($image_number_label)
        $script:Form.Controls.add($image_location_trackbar)
        $script:Form.Controls.add($no_images_label)
        $script:Form.Controls.add($wait_label)
        #$picture_box.add_paint({paint_functions})
        $Form.ShowDialog();
    }
    $script:lock = 0;
}
################################################################################
######Load Picture Directory####################################################
function load_picture_directory
{
    $failed = 0;
    if($script:lock -eq 0)
    {
        #write-host $script:settings['TARGET_DIRECTORY']
        if(Test-Path -LiteralPath $script:settings['TARGET_DIRECTORY'])
        {
            if($script:last_picture -ne $script:current_picture)
            {
                #write-host Loading Directory
                working_directory       #Build working Directory

                ################################################################################
                ######Load Exempt/Ignored Files#################################################
                if(Test-Path -literalPath ($script:output_dir + "\Ignored.csv"))
                {
                    $ignored = Get-Content -literalPath ($script:output_dir + "\Ignored.csv")
                }
                ################################################################################
                ######Load Source Directory#####################################################
                $script:file_list = Get-ChildItem -LiteralPath $script:settings['TARGET_DIRECTORY'] -File | where {$_.extension -in ".png",".jpg",".jpeg",".bmp"} | Where-Object { $ignored -notcontains $_.FullName } | Sort-Object { [regex]::Replace($_.Name, '\d+', { $args[0].Value.PadLeft(20) }) }
                $script:file_count = $file_list.count;
                if($file_count -ge 1)
                {
                    
                    $image_number_label.Text = [string]$script:current_picture + " of $script:file_count"
                    $image_location_trackbar.TickFrequency = 1
                    $image_location_trackbar.SetRange(1, $file_count)
                    $no_images_label.hide();
                    $picture_box.Show()
                    $image_location_trackbar.Show()
                    $next_button1.enabled = $true
                    $next_button2.enabled = $true
                    $previous_button1.enabled = $true
                    $previous_button2.enabled = $true
                    $caption_box.enabled = $true
                    $image_number_label.Show()
                    $dimensions_label2.text = "N/A"
                    $image_name_label2.text = "N/A"
                    $caption_box_combo.enabled = $true
                    $rotate_right_button.enabled = $true
                    $rotate_left_button.enabled = $true
                    $script:revert_button.enabled = $true
                    $script:working_dir_button.enabled = $true
                    $script:source_dir_button.enabled = $true
                    $script:wrong_dim_button.enabled = $true
                    $script:no_captions_button.enabled = $true

                    $counter = 0
                    foreach($script:image_file in $script:file_list)
                    {
                        #write-host $script:image_file.FullName
                        $counter++
                        if($counter -eq $script:current_picture)
                        {
                            #write-host Loading Picture
                            load_picture $script:image_file
                            break
                        }
                    }
                }
                else
                {
                    $failed = 1;
                }
            }#NE Same Pic
            
        }
        else
        {
            $failed = 1;
        }
        ##################################
        if($failed -eq 1)
        {
            $picture_box.Hide()
            $image_location_trackbar.Hide()
            $next_button1.enabled = $false
            $next_button2.enabled = $false
            $previous_button1.enabled = $false
            $previous_button2.enabled = $false
            $caption_box.enabled = $false
            $image_number_label.Hide()
            $dimensions_label2.text = "N/A"
            $image_name_label2.text = "N/A"
            $caption_box_combo.enabled = $false
            $rotate_right_button.enabled = $false
            $rotate_left_button.enabled = $false
            $script:revert_button.enabled = $false
            $script:working_dir_button.enabled = $false
            $script:source_dir_button.enabled = $false
            $script:wrong_dim_button.enabled = $false
            $script:no_captions_button.enabled = $false
            $no_images_label.show();

        }
    }#Script Lock
}
################################################################################
######Load Picture##############################################################
function load_picture($script:image_file)
{
    ################################################################################
    ######Load Variables
    $script:image_path           = [string]$script:settings['TARGET_DIRECTORY'] + "\$script:image_file"
    $script:image_path_working   = [string]$script:output_dir + "\$script:image_file"
    $script:caption_path         = [string]$script:settings['TARGET_DIRECTORY'] + "\" + [io.path]::GetFileNameWithoutExtension($image_path) + ".txt"
    $script:caption_path_working = [string]$script:output_dir + "\" + [io.path]::GetFileNameWithoutExtension($image_path) + ".txt"
    $c_hold = 0; #Holds Picture Base on Captions
    $s_hold = 0; #Holds Picture Based on Size

  
    ################################################################################
    ######Transfer Files to Working Directory
    if(!(Test-Path -literalpath $script:image_path_working))
    {
        Copy-Item $script:image_path $script:image_path_working
    }
    if((Test-Path -literalpath $script:caption_path) -and (!(Test-Path -literalpath $script:caption_path_working)))
    {
        Copy-Item $script:caption_path $script:caption_path_working
        $caption_box.text = Get-Content $script:caption_path_working
        $caption_box.AccessibleDescription = ""
    }
    elseif(Test-Path -literalpath $script:caption_path_working)
    {
        $caption_box.text = Get-Content $script:caption_path_working
        $caption_box.AccessibleDescription = ""
    }
    else
    {
        $caption_box.text = ""
        $caption_box.AccessibleDescription = ""
    }


    ################################################################################
    ######Check Captions
    if(($caption_box.text -eq "") -and ($script:settings['FIND_CAPTIONS'] -eq "On"))
    {
        $c_hold = 1; #No Captions Hold
    }


    ################################################################################
    ######Start Trying
    try
    {
        
        ################################################################################
        ######Load Image to Memory
        if($script:debug -eq 1) {write-host Phase 2: Loading to Memory}
        if(($script:click_active -ne 1) -or ($script:current_picture -ne $script:last_picture) -or ($script:bitmap -eq $null))
        {
            if(Test-Path variable:bitmap){$script:bitmap.Dispose()}
            [System.GC]::GetTotalMemory(‘forcefullcollection’) | out-null
            $script:picture_box.Refresh(); #Reset Canvas
             
            if(Test-path -literalpath $script:image_path_working)
            {
                $script:bitmap = [System.Drawing.Image]::Fromfile($script:image_path_working);
            }
            else
            {
                $script:bitmap = [System.Drawing.Image]::Fromfile($script:image_path);
            }
        

        
            ################################################################################
            ######Get Image Size
            if($script:debug -eq 1) {write-host Phase 3: Getting Image Size}
            $script:image_height = $script:bitmap.height
            $script:image_width = $script:bitmap.width
            $script:image_height_s = 0
            $script:image_width_s = 0
            $dim = "$image_width x $image_height"
            $dimensions_label2.Text = "$dim"

            ################################################################################
            ######Check Image Dimensions
            if($script:debug -eq 1) {write-host Phase 4: Checking Dimensions}
            if(($script:image_width -ne $script:settings['TARGET_WIDTH']) -or ($script:image_height -ne $script:settings['TARGET_HEIGHT']))
            {
                $dimensions_label2.ForeColor = "Red"
                $s_hold = 1; #Wrong Dimensions Hold
            }
            else
            {
                $dimensions_label2.ForeColor = "darkgray"
            }
        

            ################################################################################
            ######Wrong Dimension & No Caption Skipping System
            if($script:debug -eq 1) {write-host Phase 5: Skip System}
            $go_next = 0;
            if(($script:settings['FIND_DIMENSIONS'] -eq "On") -and ($script:settings['FIND_CAPTIONS'] -eq "On"))
            {
                if(($s_hold -eq 1) -or ($c_hold -eq 1))
                {
                    #Stay Here        
                }
                else
                {
                    $go_next = 1;
                }
            }
            elseif(($script:settings['FIND_DIMENSIONS'] -eq "On") -and ($s_hold -eq 0))
            {
                $go_next = 1;
            }
            elseif(($script:settings['FIND_CAPTIONS'] -eq "On") -and ($c_hold -eq 0))
            {
                $go_next = 1;
            }

        
            ################################################################################
            ######Start Skipping
            if($script:debug -eq 1) {write-host Phase 6: Start Skipping}
            if($script:target_lock -ne 1)
            {
                if($go_next -eq 1)
                {
                    $script:lock = 1;
                    $script:current_picture++;
                    $counter = 0;
                    foreach($script:image_file in $script:file_list)
                    {
                        $counter++
                        if($counter -eq $script:current_picture)
                        {
                            $script:good_up_to = $counter
                            $image_number_label.Text = [string]$script:current_picture + " of $file_count"
                            $image_location_trackbar.Value = $script:current_picture
                            load_picture $script:image_file
                            $script:lock = 0;
                            break
                        }
                    }
                }
            }
            else
            {
                $script:target_lock = 0;
            }

        
            ################################################################################
            ######Update Next/Previous Buttons
            if($script:debug -eq 1) {write-host Phase 7: Update Buttons}
            if(($script:current_picture -le 1) -or ($script:current_picture -le $script:good_up_to))
            {
                #write-host Good $script:good_up_to
                $previous_button1.enabled = $false
                $previous_button2.enabled = $false
            }
            else
            {
                $previous_button1.enabled = $true
                $previous_button2.enabled = $true
            }
            if($script:current_picture -ge $file_count)
            {
                $next_button1.enabled = $false
                $next_button2.enabled = $false
            }
            else
            {
                $next_button1.enabled = $true
                $next_button2.enabled = $true
            }

        
            ################################################################################
            ######Load Captions
            if($script:debug -eq 1) {write-host Phase 8: Load Captions}
            load_directory_captions
            find_best_captions
            


            ################################################################################
            ######Load Canvas Variables
            if($script:debug -eq 1) {write-host Phase 9: Loading Canvas Vars}
            $canvas_width = $script:picture_box.width
            $canvas_height = $script:picture_box.height
            $script:canvas_x = 0
            $script:canvas_y = 0


            ################################################################################
            ######Scale Image & Center Image on Canvas By Height
            if($script:debug -eq 1) {write-host Phase 10: Scale Image}
            if($script:image_height -ge $script:image_width)
            {
                $tries = 0
                while(($tries -lt 200))
                {
                    $canvas_height_buffer = ($canvas_height - $tries)
                    $script:scale = (($canvas_height_buffer / $script:image_height) * 100)
                    $script:image_height_s = $canvas_height_buffer
                    $script:image_width_s = (($script:image_width / 100) * $script:scale )
                    [int]$script:canvas_y = (($canvas_height / 2) - ($script:image_height_s / 2))
                    [int]$script:canvas_x = (($canvas_width / 2) - ($script:image_width_s / 2))
                    determine_crosshair_size
                    if(($canvas_height -ge $script:image_height_s) -and ($canvas_width -ge $script:image_width_s))
                    {
                        break;
                    }
                    $tries++
                }#while
            }
            ################################################################################
            ######Scale Image & Center Image on Canvas By Width
            else
            {
                $tries = 0
                while(($tries -lt 200))
                {
                    $canvas_width_buffer = ($canvas_width - $tries)
                    $script:scale = (($canvas_width_buffer / $script:image_width) * 100)
                    $script:image_width_s = $canvas_width_buffer
                    $script:image_height_s = (($script:image_height / 100) * $scale)
                    [int]$script:canvas_y = (($canvas_height / 2) - ($script:image_height_s / 2))
                    [int]$script:canvas_x = (($canvas_width / 2) - ($script:image_width_s / 2))
                    determine_crosshair_size
                    if(($canvas_height -ge $script:image_height_s) -and ($canvas_width -ge $script:image_width_s))
                    {
                        break;
                    }

                    $tries++
                }#while
            }

            if($script:debug -eq 1) {write-host Phase 11: Place Image on Canvas}
            ################################################################################
            ######Store Boundry Box Variables
            $script:image_location_point_corner1x = $script:canvas_x
            $script:image_location_point_corner1y = $script:canvas_y
            $script:image_location_point_corner2x = $script:canvas_x + $script:image_width_s -1
            $script:image_location_point_corner2y = $script:canvas_y + $script:image_height_s -1

        }#Not Last Image
        ################################################################################
        ######Place Image on Canvas
        $script:FormGraphics.DrawImage($script:bitmap, $script:canvas_x, $script:canvas_y, $script:image_width_s, $script:image_height_s)
        $image_name_label2.Text = [io.path]::GetFileNameWithoutExtension($image_path)
        
        $script:last_picture = $script:current_picture  
    }
    catch
    {
        write-host FAILED
    }
    
}
################################################################################
######Prompt for Folder#########################################################
function prompt_for_folder()
{  
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms")|Out-Null

    $foldername = New-Object System.Windows.Forms.FolderBrowserDialog
    $foldername.Description = "Select a folder"
    $foldername.rootfolder = "MyComputer"
    if(($script:settings['TARGET_DIRECTORY'] -ne "") -and (Test-Path -literalpath $script:settings['TARGET_DIRECTORY']))
    {
        $foldername.SelectedPath = $script:settings['TARGET_DIRECTORY']
    }
    if($foldername.ShowDialog() -eq "OK")
    {
        $folder += $foldername.SelectedPath
    }
    $script:settings['TARGET_DIRECTORY'] = $folder
}
################################################################################
######Paint Functions###########################################################
function paint_functions
{
    #write-host Painting
}
################################################################################
######Idle Timer################################################################
Function Idle_Timer
{
    #write-host $Script:Timer.Interval
    $Script:CountDown = $Script:CountDown + 1;
    
    ################################################################################
    ######Form Resize###############################################################
    $script:lock = 1;
    if(($script:Form.width -ne $script:form_width) -or ($script:Form.height -ne $script:form_height) -and ($script:started -ne 0))
    { 
        $script:form_width = $script:Form.width
        $script:form_height = $script:Form.height

        write-host $script:form_width x $script:form_height
        $script:last_picture = 0;
        main          
    }
    $script:lock = 0;


    ################################################################################
    ######Load Pictures#############################################################     
    load_picture_directory


    ################################################################################
    ######Click Activations#########################################################
    if($script:click_active -eq 1)
    {
        $script:target_lock = 1
        ##############################
        ###########Get Cursor
        if($script:zoom -ne 1)
        {
            $x = ((([System.Windows.Forms.Cursor]::Position.X) - ($script:Form.Location.X)) - $script:picture_box.location.x)
            $y = ((([System.Windows.Forms.Cursor]::Position.Y) - ($script:Form.Location.Y)) - $script:picture_box.location.y - 30)
            $script:last_click_x = $x ##Save For Zoom
            $script:last_click_y = $y ##Save For Zoom
        }

        ##############################
        ###########Reset Canvas Image
        [string]$image = [io.path]::GetFileNameWithoutExtension($image_path) + [System.IO.Path]::GetExtension($image_path)
        load_picture $image #Reset Image
        
        ##############################
        ###########Calculate Cursor Offset
        $base_x = ((2 / ($script:crosshair_size_x * 5)) * 100)
        $base_y = ((2 / ($script:crosshair_size_y * 5)) * 100)
        $x = (($script:crosshair_size_x / (2 + $base_x)) + $x) #Cursor Offset
        $y = (($script:crosshair_size_y / (2 + $base_y)) + $y) #Cursor Offset

        ##############################
        ###########Draw Crosshair
        $myBrush = new-object Drawing.SolidBrush yellow
    
        #Rectangle 
        $c1x = ($x - $script:crosshair_size_x)
        $c2x = $script:crosshair_size_x
        $c1y = ($y - $script:crosshair_size_y)
        $c2y = $script:crosshair_size_y

        #Horizontal Line
        $h1x = $x
        $h2x = ($x - $script:crosshair_size_x)
        $h1y = ($y - ($script:crosshair_size_y / 2))
        $h2y = ($y - ($script:crosshair_size_y / 2))

        #Vertical Line
        $v1x = ($x - ($script:crosshair_size_x / 2))
        $v2x = ($x - ($script:crosshair_size_x / 2))
        $v1y = $y  - 2
        $v2y = ($y - $script:crosshair_size_y)
        

        ##############################
        ###########Horizontal Boundries
        if($c1x -le $script:image_location_point_corner1x)
        {
            $myBrush = new-object Drawing.SolidBrush Red
            $c1x = $script:image_location_point_corner1x
            $h1x = $script:image_location_point_corner1x
            $h2x = ($h1x + $script:crosshair_size_x)
            $v1x = $script:image_location_point_corner1x + ($script:crosshair_size_x / 2)
            $v2x = $script:image_location_point_corner1x + ($script:crosshair_size_x / 2)
        }
        elseif($x -ge $script:image_location_point_corner2x)
        {
            $myBrush = new-object Drawing.SolidBrush Red
            $c1x = ($script:image_location_point_corner2x - $script:crosshair_size_x) 
            $h1x = ($script:image_location_point_corner2x - $script:crosshair_size_x)
            $h2x = ($h1x + $script:crosshair_size_x)
            $v1x = $script:image_location_point_corner2x - ($script:crosshair_size_x / 2)
            $v2x = $script:image_location_point_corner2x - ($script:crosshair_size_x / 2)
        }
        ##############################
        ###########Vertical Boundries
        if(($y - $script:crosshair_size_y) -le $script:image_location_point_corner1y)
        {
            $myBrush = new-object Drawing.SolidBrush Red
            $c1y = $script:image_location_point_corner1y
            $h1y = $script:image_location_point_corner1y + ($script:crosshair_size_y / 2)
            $h2y = $h1y
            $v1y = $script:image_location_point_corner1y + $script:crosshair_size_y
            $v2y = $script:image_location_point_corner1y
        }
        elseif(($y -ge $script:image_location_point_corner2y))
        {
            $myBrush = new-object Drawing.SolidBrush Red
            $c1y = ($script:image_location_point_corner2y - $script:crosshair_size_y)
            $h2y = $script:image_location_point_corner2y - ($script:crosshair_size_y / 2)
            $h1y = $h2y
            $v1y = $script:image_location_point_corner2y 
            $v2y = $script:image_location_point_corner2y - $script:crosshair_size_y
        }

        ##############################
        ###########Draw Crosshair
        $rectangle = new-object Drawing.Rectangle  $c1x, $c1y, $c2x, $c2y
        $formGraphics.DrawRectangle($myBrush, $rectangle)   
        $formGraphics.DrawLine($myBrush, $h1x, $h1y, $h2x,$h2y) #Horizontal Line 
        $formGraphics.DrawLine($myBrush, $v1x, $v1y ,$v2x,$v2y) #Vertical


        ##############################
        ###########Translate Scope to Actual  
        $diff_x = (($script:picture_box.width - $script:image_width_s) / 2)
        $diff_y = (($script:picture_box.height - $script:image_height_s) / 2)
        $c1x = $c1x - $diff_x
        $c1y = $c1y - $diff_y
        $script:crosshair_x1 = (($script:image_width / 100) * (($c1x / $script:image_width_s ) * 100))
        $script:crosshair_x2 = (($script:image_width / 100) * (($c2x / $script:image_width_s ) * 100)) #+ $script:crosshair_x1
        $script:crosshair_y1 = (($script:image_height / 100) * (($c1y / $script:image_height_s ) * 100))
        $script:crosshair_y2 = (($script:image_height / 100) * (($c2y / $script:image_height_s ) * 100)) #+ $script:crosshair_y1
    }
    ################################################################################
    ######Idle Timer Caption Box####################################################
    if(($script:lock -ne 1) -and ($caption_box.text -ne $caption_box.AccessibleDescription))
    {
        if(($script:image_path_working -ne $null) -and (Test-Path -LiteralPath $script:image_path_working))
        {
            $caption_box.AccessibleDescription = $caption_box.text
            Set-Content $script:caption_path_working $caption_box.text
            if(!($script:caption_list.containsvalue($caption_box.text)))
            {
                $script:caption_list[$script:current_picture] = $caption_box.text
                
            }
            find_best_captions
        }
    }
    ################################################################################
    ######Idle Timer Wait Actions###################################################
    if(($script:wait -ne 0) -and ($Script:CountDown -gt $script:wait))
    {
        $script:wait = 0;
        if($script:wait_mode -eq "Next")
        {
            $script:current_picture++;
            $image_location_trackbar.Value = $script:current_picture
            $wait_label.hide();
        }
        elseif($script:wait_mode -eq "Prev")
        {
            $script:current_picture--;
            $image_location_trackbar.Value = $script:current_picture
            $wait_label.hide();
        }
        elseif($script:wait_mode -eq "Trackbar")
        {
            $Script:Timer.Interval = 1000
            if($caption_box.text -ne $caption_box.AccessibleDescription)
            {
                $caption_box.AccessibleDescription = $caption_box.text
                Set-Content $script:caption_path_working $caption_box.text
            }
            $script:current_picture = $image_location_trackbar.Value
            $image_number_label.Text = [string]$script:current_picture + " of $script:file_count"
        }
    }
    elseif(($script:wait -ne 0) -and ($script:wait_mode -match "Next|Prev"))
    {
        $wait_label.text = (($script:wait - $Script:CountDown) + 1)
        $wait_label.show()
    }
    

}
################################################################################
######Process Crosshair#########################################################
function process_crosshair
{
    $script:target_lock = 1
    #write-host X1 $script:crosshair_x1
    #write-host Y1 $script:crosshair_y1
    #write-host X2 $script:crosshair_x2
    #write-host Y2 $script:crosshair_y2

    ###Capture Scope to Image
    $units    = [System.Drawing.GraphicsUnit]::Pixel
    $bmp      = new-object System.Drawing.Bitmap ([int]$script:crosshair_x2,[int]$script:crosshair_y2)
    $graphics = [System.Drawing.Graphics]::FromImage($bmp)
    $destRect = new-object Drawing.Rectangle 0, 0, $script:crosshair_x2, $script:crosshair_y2
    $srcRect  = new-object Drawing.Rectangle $script:crosshair_x1, $script:crosshair_y1, $script:crosshair_x2, $script:crosshair_y2
    $graphics.DrawImage($script:bitmap, $destRect, $srcRect, $units)
    $script:bitmap.Dispose()
    [System.GC]::GetTotalMemory(‘forcefullcollection’) | out-null

    ####Resize Image to Settings
    $bmp_resize      = new-object System.Drawing.Bitmap ([int]$script:settings['TARGET_WIDTH'],[int]$script:settings['TARGET_HEIGHT'])
    $graphics = [System.Drawing.Graphics]::FromImage($bmp_resize)
    $srcRect   = new-object Drawing.Rectangle 0, 0, $script:crosshair_x2, $script:crosshair_y2
    $destRect  = new-object Drawing.Rectangle 0, 0, $script:settings['TARGET_WIDTH'], $script:settings['TARGET_HEIGHT']
    $graphics.DrawImage($bmp, $destRect, $srcRect, $units)

    $bmp_resize.Save($script:image_path_working)  
    $picture_box.refresh();
    $script:last_picture = 0;
    



}
################################################################################
######Building Working Directory################################################
function working_directory
{
    if(($script:output_dir -eq "") -or (!(Test-Path -LiteralPath $script:output_dir)))
    {
        $working_dir = $script:settings['TARGET_DIRECTORY'].split('\\|/|:',[System.StringSplitOptions]::RemoveEmptyEntries)
        if($working_dir[-3])
        {
            $working_dir = $working_dir[-3] + "-" + $working_dir[-2] + "-" + $working_dir[-1]
        }
        elseif($working_dir[-2])
        {
            $working_dir = $working_dir[-2] + "-" + $working_dir[-1]
        }
        else
        {
            $working_dir = $working_dir[-1]
        }
        $script:output_dir = $script:user_output_dir + "\$working_dir"
        if(!(Test-Path $script:output_dir))
        {
            New-Item -Path $script:output_dir -ItemType Directory | Out-Null
        }
    }
}
################################################################################
######Load Directory Captions###################################################
function load_directory_captions
{
    if($script:caption_list.Count -eq 0)
    {
        $script:output_dir
        Get-ChildItem -LiteralPath $script:output_dir -Filter *.txt | Foreach-Object {
            
            $string = Get-Content -LiteralPath $_.Fullname -First 1
            $name = [System.IO.Path]::GetFileNameWithoutExtension($_.Fullname)
            if($string.Length -le 1500)
            {
                if(!($script:caption_list.containsvalue($string)))
                {
                    #write-host $name = $string
                    $script:caption_list[$name] = $string
                }
            }
        }
    }

}
################################################################################
######Resize Image##############################################################
function resize_image
{
    ####Resize Image to Settings
    $units    = [System.Drawing.GraphicsUnit]::Pixel
    $bmp_resize      = new-object System.Drawing.Bitmap ([int]$script:settings['TARGET_WIDTH'],[int]$script:settings['TARGET_HEIGHT'])
    $graphics = [System.Drawing.Graphics]::FromImage($bmp_resize)
    $srcRect   = new-object Drawing.Rectangle 0, 0, $script:image_width, $script:image_height
    $destRect  = new-object Drawing.Rectangle 0, 0, $script:settings['TARGET_WIDTH'], $script:settings['TARGET_HEIGHT']
    $graphics.DrawImage($script:bitmap, $destRect, $srcRect, $units)
    $script:bitmap.Dispose()
    [System.GC]::GetTotalMemory(‘forcefullcollection’) | out-null
    $bmp_resize.Save($script:image_path_working)
    $script:last_picture = 0;
    $script:wait = $Script:CountDown + 2
}
################################################################################
######Rotate Image##############################################################
function rotate_image($direction)
{
    
    ####Resize Image to Settings
    if($direction -eq "Right")
    {
        $script:bitmap.rotateflip("Rotate90FlipNone")
        $script:bitmap.save($script:image_path_working)
    }
    else
    {
        $script:bitmap.rotateflip("Rotate270FlipNone")
        $script:bitmap.save($script:image_path_working)
    }
    $script:last_picture = 0;

}
################################################################################
######Determine Cross Hair Size#################################################
function determine_crosshair_size($mode)
{
    if($mode -eq "increase")
    {
        $increase = 20;
        while($increase -ge 0)
        {
            if((($script:crosshair_size_x + $increase) -le $script:image_width_s) -and (($script:crosshair_size_y + $increase) -le $script:image_height_s) )
            {
                $script:crosshair_size_x = $script:crosshair_size_x + $increase
                $script:crosshair_size_y = $script:crosshair_size_y + $increase
                break;
            }
            $increase--
        }
    }
    if($mode -eq "decrease")
    {
        $decrease = 20;
        while($decrease -ge 0)
        {
            if((($script:crosshair_size_x - $decrease) -ge 30) -and (($script:crosshair_size_y - $decrease) -ge 30) )
            {
                $script:crosshair_size_x = $script:crosshair_size_x - $decrease
                $script:crosshair_size_y = $script:crosshair_size_y - $decrease
                break;
            }
            $decrease--
        }
    }
    if(($script:last_picture -ne $script:current_picture) -or ($mode -eq "image"))
    {
        $max_size = 0
        if($script:image_width_s -ge $script:image_height_s)
        {
            $max_size = $script:image_height_s
        }
        else
        {
            $max_size = $script:image_width_s
        }
        $scaled = 10.2;
        while($scaled -ge 1)
        {
            
            $script:crosshair_size_x_buffer = (($script:settings['TARGET_WIDTH'] / ($max_size / $scaled)) * $script:scale)
            $script:crosshair_size_y_buffer = (($script:settings['TARGET_HEIGHT'] / ($max_size / $scaled)) * $script:scale)
            if(($script:crosshair_size_x_buffer -le ($script:image_width_s - 1)) -and ($script:crosshair_size_y_buffer -le ($script:image_height_s - 1)))
            {  
                $script:crosshair_size_x = $script:crosshair_size_x_buffer
                $script:crosshair_size_y = $script:crosshair_size_y_buffer
                break;
            }
            $scaled = $scaled - 0.2 
        }
    }
    $mode = "";
}
################################################################################
######Find Best Captions########################################################
function find_best_captions
{
    $script:rack_and_stack = @{};
    $current_caption = $caption_box.text
    $current_words = $current_caption.ToLower() -replace "[^a-z0-9]| | ",' '
            

    $current_caption_modified = $current_caption.ToLower() -replace "[^a-z0-9]| | ",' '
    $current_caption_modified_wordsplit = $current_caption_modified -split ' ';
    $current_caption_modified_wordsplit = ($current_caption_modified_wordsplit | sort length -desc | select -first 6)

        
    foreach($caption in $script:caption_list.getEnumerator())
    {
        $caption = $caption.value
        $match_score = 0;
        $match_caption = $caption.ToLower() -replace "[^a-z0-9]| | ",' '
        foreach($word in $current_caption_modified_wordsplit)
        {
            if($match_caption -match "$word")
            {
                $match_score = $match_score + $word.length;
            }
        }
        if($current_caption -ne $caption)
        {
            if(($script:rack_and_stack.Get_Count()) -le 14)
            {
                if(!($script:rack_and_stack.Contains($caption)))
                {
                            
                    $script:rack_and_stack.Add("$caption",$match_score);
                } 
            }
            else
            { 
                if(!($script:rack_and_stack.Contains("$caption")))
                {
                    foreach($scored_caption in $script:rack_and_stack.getEnumerator() | Sort Value -Descending) 
                    {
                        $scored_caption1 = $scored_caption.key
                        if($match_score -gt $scored_caption.value)
                        {
                            $script:rack_and_stack.Remove("$scored_caption1")
                            $script:rack_and_stack.Add("$caption",$match_score);
                            break;
                        }
                    }

                }
            }
        }
    }
    $script:lock = 1;
    $script:caption_box_combo.Items.Clear();
    $counter = 0;
    foreach($scored_caption in $script:rack_and_stack.getEnumerator() | Sort Value -Descending) 
    {
        $script:caption_box_combo.items.add($scored_caption.key)
        $counter++;
        if($counter -eq 1)
        {
            $script:caption_box_combo.selectedItem = $scored_caption.key
        }
        #write-host $scored_caption.value = $scored_caption.key
        
    }
    $script:lock = 0;
}
################################################################################
######Load Settings#############################################################
function load_settings
{
    if(Test-Path "$dir\Settings.csv")
    {
        $line_count = 0;
        $reader = [System.IO.File]::OpenText("$dir\Settings.csv")
        while($null -ne ($line = $reader.ReadLine()))
        {
            $line_count++;
            if($line_count -ne 1)
            {
                ($key,$value) = $line -split ',',2
                if(!($script:settings.containskey($key)))
                {
                    $script:settings.Add($key,$value);
                }
            } 
        }
        $reader.close();
    }
    if($script:settings['TARGET_DIRECTORY'] -eq $null)
    {
        $script:settings['TARGET_DIRECTORY'] = "Browse or Enter a file path"
    }
}
################################################################################
#####Update Settings############################################################
function update_settings
{
    if($script:settings.count -ne 0)
    {
        if(Test-Path "$dir\Buffer_Settings.csv")
        {
            Remove-Item -literalpath "$dir\Buffer_Settings.csv"
        }
        $buffer_settings = new-object system.IO.StreamWriter("$dir\Buffer_Settings.csv",$true)
        $buffer_settings.write("PROPERTY,VALUE`r`n");
        foreach($setting in $script:settings.getEnumerator() | Sort key)                  #Loop through Input Entries
        {
                $setting_key = $setting.Key                                               
                $setting_value = $setting.Value
                $buffer_settings.write("$setting_key,$setting_value`r`n");
        }
        $buffer_settings.close();
        if(test-path -LiteralPath "$dir\Buffer_Settings.csv")
        {
            if(Test-Path -LiteralPath "$dir\Settings.csv")
            {
                Remove-Item -LiteralPath "$dir\Settings.csv"
            }
            Rename-Item -LiteralPath "$dir\Buffer_Settings.csv" "$dir\Settings.csv"
        }
    } 

}
################################################################################
######Initial Checks############################################################
function initial_checks
{
    if(!(Test-Path -LiteralPath "$dir\Settings.csv"))
    {
        $settings_writer = new-object system.IO.StreamWriter("$dir\Settings.csv",$true)
        $settings_writer.write("TARGET_DIRECTORY,Browse or Enter a file path`r`n");
        $settings_writer.write("TARGET_HEIGHT,512`r`n");
        $settings_writer.write("TARGET_WIDTH,512`r`n");
        $settings_writer.write("FIND_CAPTIONS,Off`r`n");
        $settings_writer.write("FIND_DIMENSIONS,Off`r`n");
        $settings_writer.close();
    }
}
################################################################################
######Main Sequence#############################################################
initial_checks
load_settings
main
