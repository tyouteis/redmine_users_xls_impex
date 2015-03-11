require 'spreadsheet'

class UserXlsImpexColumn
  attr_accessor :name
  include Redmine::I18n

  def initialize(name,global_field=true)
    self.name=name
    @caption_str_pfx=(global_field==true ? "field_#{name}": "label_users_xlsie_field_#{name}" )
  end

  def caption
    l(@caption_str_pfx)
  end

  def value(user)
    user.send(name)
  end
end

class UserXlsImpexCustomFieldColumn < UserXlsImpexColumn
  def initialize(custom_field)
    self.name = "cf_#{custom_field.id}".to_sym
    @cf=custom_field
  end

  def caption
    @cf.name
  end

  def value(user)
    cv = user.custom_values.detect {|v| v.custom_field_id == @cf.id}
    cv && @cf.cast_value(cv.value)
  end

  def custom_field
    @cf
  end
end

class IncludeInImportUserXlsImpexColumn < UserXlsImpexColumn
  def value(user)
    0
  end
end

class IsAdminUserXlsImpexColumn < UserXlsImpexColumn
  def value(user)
    user.admin? ? 1 : 0
  end
end

class UsersXLSIARCondition
  attr_reader :conditions

  def initialize(condition=nil)
    @conditions = ['1=1']
    add(condition) if condition
  end

  def add(condition)
    if condition.is_a?(Array)
      @conditions.first << " AND (#{condition.first})"
      @conditions += condition[1..-1]
    elsif condition.is_a?(String)
      @conditions.first << " AND (#{condition})"
    else
      raise "Unsupported #{condition.class} condition: #{condition}"
    end
    self
  end

  def <<(condition)
    add(condition)
  end
end

class UsersXlsImpexController < ApplicationController
  unloadable

  helper :sort
  include SortHelper
  helper :custom_fields
  include CustomFieldsHelper

  before_filter :require_admin

  def export_start
   if request.post?

      sort_init 'id', 'asc'
      sort_update %w(id login firstname lastname mail admin created_on last_login_on)

      @export_status = params[:export_status] ? params[:export_status].to_i : 1
      c = UsersXLSIARCondition.new(@export_status == 0 ? "status <> 0" : ["status = ?", @export_status])

      unless params[:name].blank?
        name = "%#{params[:name].strip.downcase}%"
        c << ["LOWER(login) LIKE ? OR LOWER(firstname) LIKE ? OR LOWER(lastname) LIKE ? OR LOWER(mail) LIKE ? OR LOWER(vvk_dept) LIKE ?", name, name, name, name, name]
      end

      @users =  User.find(:all,:order => sort_clause,
                          :conditions => c.conditions)

      export_name = 'users_export_data.xls'
      send_data(users_to_xls2(@users), :type => :xls, :filename => export_name)
    end
  end

  def import_start
    if !params[:to_users].blank?
      remove_temp_file
      reset_session_data
      redirect_to :controller => 'users', :action => 'index'
      return
    end
    if !params[:step_back].blank?
      redirect_to :action => 'import_results', :create_imported_users => true, :send_user_info => params[:send_user_info]
      return
    end
    if !params[:step_back_2].blank?
      redirect_to :action => 'import'
      return
    end
    remove_temp_file
    reset_session_data
  end

  def prepare_import
    if params[:xls_file].blank?
        redirect_to :action => 'import_start'
        return
    end
    download_xls_file
    session[:users_xlsie_file]=@xls_book_name
    load_xls_file
  end

  def import
    if !params[:back_to_index].blank?
      redirect_to :action => 'import_start'
      return
    end
    retrieve_prepare_attrs
    @xls_sheet_num=session[:users_xlsie_sheet]
    @xls_book_name=session[:users_xlsie_file]

    load_xls_file
    retrieve_columns(:import)
    fill_helper_columns
  end

  def import_preview
    if !params[:mode].blank?
      @text=create_filters_wiki_help
      render :partial => 'common/preview'
      return
    end

    retrieve_import_result_params
    @xls_book_name=session[:users_xlsie_file]
    logger.info("[XLSI] Book name from session store is #{@xls_book_name}")
    @xls_sheet_num=session[:users_xlsie_sheet]
    xls_columns_map = session[:users_xlsie_columns_map]

    load_xls_file
    retrieve_columns(:import)

    @xls_columns = []
    @formatted_values = []

    sheet1=@xls_book.worksheet(@xls_sheet_num)

    sheet1.row(0).each_with_index do |r,idx|
      next if r.to_s[0..0] == '!'
      column_name=xls_columns_map[idx.to_s]
      if column_name!=''
        c=@user_columns.detect do |ic|
          true if ic.name == column_name.to_sym
        end
        @xls_columns << [ r.to_s, c!=nil ? c.caption : '' ]
      end
    end

    header_row=sheet1.row(0)
    (1..10).each do |idx|
      row_values = []
      row=sheet1.row(idx)
      break if row.size==0
      row.each_with_index do |r,idx2|
        next if header_row[idx2].to_s[0..0] == '!'

        if xls_columns_map[idx2.to_s]!=''
          if row[idx2] == nil
            row_values << ''
            next
          end

          if row[idx2].is_a?(Spreadsheet::Formula)
# hack to get rid of formula class
            row_preformatted_value=row[idx2].value
            row[idx2]=row_preformatted_value
          end

          if row[idx2] == nil
            row_values << ''
            next
          end

          value=format_row_value_before(row,idx2)
          if value == nil
            row_values << ''
          else
            row_values << value.to_s
          end
        end
      end
      @formatted_values << row_values
    end

    render :partial => 'import_preview'
  end

  def import_results
    unless params[:back_to_import].blank?
      redirect_to :action => 'import_start'
      return
    end
    retrieve_import_result_params
    @validate_only = params[:create_imported_users].blank? ? true : false
    @xls_book_name=session[:users_xlsie_file]
    @xls_sheet_num=session[:users_xlsie_sheet]
    xls_columns_map = session[:users_xlsie_columns_map]
    user_columns_map = xls_columns_map.invert
    @update_users_mode = session[:users_xlsie_update_only].to_i
    send_user_info = params[:send_user_info].blank?  ? false : true

    retrieve_columns(:import)
    load_xls_file

    sheet1=@xls_book.worksheet(@xls_sheet_num)

    @rows_skipped = 0
    # each element is an array of users which are updated
    @users_updated = []
    # each element is an array of users which are duplicated
    @users_duplicated = []

    # Each element of the arrays below has following format
    # [:row] - xls row for this record
    # [:user] - user object
    # [:validation_errors] - array of [:attr_name,:message]
    @users_saved = Array.new
    @users_failed = Array.new

    curr_row=0

    sheet1.each 1 do |row|

      curr_row += 1

      unless user_columns_map['include_in_import'].blank?
        row_idx = user_columns_map['include_in_import'].to_i
        if row[row_idx].is_a?(Spreadsheet::Formula)
# hack to get rid of formula class
          row_preformatted_value=row[row_idx].value
          row[row_idx]=row_preformatted_value
        end
        if row[row_idx] == nil || row[row_idx].to_i != 1
# if row[row_idx] == nil || row[row_idx] == '' || row[row_idx] ==' ' || row[row_idx].to_i == 0
          @rows_skipped += 1
          next
        end
      end

      user_params = Hash.new

      @user_columns.each do |c|

        next if user_columns_map[c.name.to_s].blank?

        row_idx=user_columns_map[c.name.to_s].to_i
        next if row[row_idx] == nil

        if row[row_idx].is_a?(Spreadsheet::Formula)
# hack to get rid of formula class
          row_preformatted_value=row[row_idx].value
          row[row_idx]=row_preformatted_value
        end

        next if row[row_idx] == nil

        row_value = format_row_value_before(row,row_idx)

        if c.is_a?(UserXlsImpexCustomFieldColumn)
          user_params['custom_field_values'] = Hash.new if user_params['custom_field_values'].blank?
          user_params['custom_field_values']["#{c.custom_field.id}"]=row_value.to_s
# next column
          next
        end

        user_params[c.name] = case c.name
          when :id
            row_value.to_i
        else
          row_value.to_s
        end
      end

      user_params[:password_confirmation]=user_params[:password] unless user_params[:password].blank?

      user=User.new(:language => Setting.default_language, :mail_notification => Setting.default_notification_option)
      if @update_users_mode != 0
        found_user=nil
        unless user_params[:id].blank?
          found_user = User.find_by_id(user_params[:id])
        else
          found_user = User.find_by_login(user_params[:login]) unless user_params[:login].blank?
        end
        if found_user != nil
          @users_updated << found_user
          user=found_user
        else
          if @update_users_mode == 2
            # should update only
            @rows_skipped += 1
            next
          end
        end
      end

      # should keep old custom values in existing user as well
      unless user_params['custom_field_values'].blank?
        unless user.new_record?
          user.custom_field_values.each do |v|
            if user_params['custom_field_values']["#{v.custom_field.id}"].blank?
              user_params['custom_field_values']["#{v.custom_field.id}"]=v.value if v.value != nil
            end
          end
        end
        user.custom_field_values=user_params['custom_field_values']
      end

      user_params.each_pair do |v_name, value|
        next if [:custom_field_values,:id].include?(v_name.to_sym)

        user.send("#{v_name.to_s}=",value) if user.respond_to?("#{v_name.to_s}=")
      end

      save_result = (@validate_only == true ? user.valid? : user.save)
      unless save_result
        @users_updated.delete(user.id)
        error_details = []
        #user.custom_field_values.each do |cv|
        # cv.errors.each do |attr,msg|
        #   error_details << { :attr_name => cv.custom_field.name, :message => attr+' '+msg }
        # end
        #end
        user.errors.each do |attr,msg|
          error_details << { :attr_name => attr, :message => msg }
        end
        @users_failed << { :row => curr_row, :user => user, :validation_errors => error_details }
      else
        Mailer.deliver_account_information(user, user_params[:password]) if send_user_info && !@validate_only
        @users_saved << { :row => curr_row, :user => user }
      end
    end
    unless @validate_only
      remove_temp_file
      reset_session_data
    end

# list of columns for error output
    if @users_failed.count > 0
      @xls_columns = []
      @formatted_values = []

      sheet1.row(0).each_with_index do |r,idx|
        next if r.to_s[0..0]=='!'

        column_name=xls_columns_map[idx.to_s]
        if column_name!=''
          c=@user_columns.detect do |ic|
            true if ic.name == column_name.to_sym
          end
          @xls_columns << [ r.to_s, c!=nil ? c.caption : '', c.name, idx ]
        end
      end

      @users_failed.each do |ic|
        row_values = []
        @xls_columns.each do |xc|
          ret_str = ''
          ret_str << sheet1.row(ic[:row])[xc[3]].to_s
          row_values << ret_str
        end
        @formatted_values << row_values
      end
    end

  end

private

  def download_xls_file
    tmpfile=create_temp_file
    #tmpfile=Tempfile.new("xlsi")
    file=params[:xls_file]
    file.binmode if file.respond_to?(:binmode)
    tmpfile.binmode
    tmpfile.write(file.read)
    tmpfile.close
    @xls_book_name=File.basename(tmpfile.path)
    logger.info("[XLSI] XLS file downloaded: #{@xls_book_name}, #{tmpfile.path}")
  end

  def remove_temp_file
    @xls_book_name=session[:users_xlsie_file] unless @xls_book_name
    if @xls_book_name
      logger.info("[XLSI] Removing temp file with name #{Rails.root.join('tmp',@xls_book_name)}")
      begin
        File.unlink(Rails.root.join('tmp',@xls_book_name))
      rescue
        logger.info("[XLSI] Temp file with name #{@xls_book_name} not found")
      end
      session.delete(:users_xlsie_file)
    else
      logger.info("[XLSI] @xls_book_name not defined")
    end
  end

  def load_xls_file
    Spreadsheet.client_encoding = 'UTF-8'
    xls_stream = StringIO.new('')
    tempfile=File.new(Rails.root.join('tmp',@xls_book_name), 'rb')
    unless tempfile
      logger.info("[XLSI] Trying to load non-existing file: #{@xls_book_name}")
    else
      xls_stream.write(tempfile.read)
      tempfile.close
      @xls_book = Spreadsheet.open xls_stream
      logger.info("[XLSI] Loading file #{@xls_book_name}")
    end
  end

  def retrieve_columns(mode)
    @user_columns = []

    @user_columns << IncludeInImportUserXlsImpexColumn.new(:include_in_import,false)
    @user_columns << UserXlsImpexColumn.new(:id,false)
    @user_columns << UserXlsImpexColumn.new(:login)
    @user_columns << UserXlsImpexColumn.new(:firstname)
    @user_columns << UserXlsImpexColumn.new(:lastname)
    @user_columns << UserXlsImpexColumn.new(:mail)
    @user_columns << UserXlsImpexColumn.new(:mail_notification)
    @user_columns << UserXlsImpexColumn.new(:language)
    @user_columns << UserXlsImpexColumn.new(:department) if User.new.respond_to?(:department)
    @user_columns << UserXlsImpexColumn.new(:description) if User.new.respond_to?(:description)
    if mode == :export
      @user_columns << IsAdminUserXlsImpexColumn.new(:admin)
      @user_columns << UserXlsImpexColumn.new(:identity_url)
      @user_columns << UserXlsImpexColumn.new(:type)
      @user_columns << UserXlsImpexColumn.new(:status)
      @user_columns << UserXlsImpexColumn.new(:auth_source_id,false)
      @user_columns << UserXlsImpexColumn.new(:last_login_on)
      @user_columns << UserXlsImpexColumn.new(:created_on)
      @user_columns << UserXlsImpexColumn.new(:updated_on)
      @user_columns << UserXlsImpexColumn.new(:hashed_password,false)
    else
      @user_columns << UserXlsImpexColumn.new(:password)
    end

    cf_by_type=CustomField.find(:all).group_by { |f| f.class.name}
    if cf_by_type && cf_by_type['UserCustomField']
      cf_by_type['UserCustomField'].each do |cf|
       @user_columns << UserXlsImpexCustomFieldColumn.new(cf)
      end
   end

  end

  def fill_helper_columns
    @xls_columns = []

    sheet1=@xls_book.worksheet(@xls_sheet_num)

    sheet1.row(0).each_with_index do |r,idx|
      @xls_columns << [ r.to_s, idx.to_s ] unless r.to_s[0..0]=='!'
    end
  end

  def format_row_value_before(row,row_idx)
    return row[row_idx]
  end

  def create_filters_wiki_help
    ret_str=''
    ret_str << "h1. Help will be here someday\n\n"
    ret_str << "| *Name* | *ID* | *Description* |\n"
    ret_str << "\n"

    return ret_str
  end

  def retrieve_prepare_attrs
    session[:users_xlsie_sheet] = params[:xls_sheets].to_i unless params[:xls_sheets].blank?
  end

  def retrieve_import_result_params
    session[:users_xlsie_columns_map] = params[:xls_columns_map] unless params[:xls_columns_map].blank?
    session[:users_xlsie_update_only] = params[:update_only] unless params[:update_only].blank?
  end

  def reset_session_data
    session.delete(:users_xlsie_columns_map)
    session.delete(:users_xlsie_update_only)
    session.delete(:users_xlsie_file)
    session.delete(:users_xlsie_sheet)
  end

  def update_column_width(old_w,new_w)
    return (old_w > new_w ? old_w : new_w)
  end

  def create_temp_file
    timestamp = DateTime.now.strftime("%y%m%d%H%M%S")
    ascii = '_uxlsie.xls'
    while File.exist?(Rails.root.join('tmp', "#{timestamp}_#{ascii}"))
      timestamp.succ!
    end
    return File.new(Rails.root.join('tmp', "#{timestamp}_#{ascii}"),"w+")
  end

# export part
# options are
# :group - group by type
  def users_to_xls2(users,options = {})

    Spreadsheet.client_encoding = 'UTF-8'

    options.default=false
    group_by_type = options[:group]

    book = Spreadsheet::Workbook.new

    retrieve_columns(:export)

    sheet1 = nil
    type = false
    columns_width = []
    idx = 0
# xls rows
    users.each do |user|
      if group_by_type
        new_type=user.type
        if new_type != type
          type = new_type
          update_sheet_formatting(sheet1,columns_width) if sheet1
          sheet1 = book.create_worksheet(:name => (type.blank? ? l(:label_none) : pretty_xls_tab_name(type.to_s)))
          columns_width=init_header_columns(sheet1,@user_columns)
          idx = 0
        end
      else
        if sheet1 == nil
          sheet1 = book.create_worksheet(:name => l(:label_user_plural))
          columns_width=init_header_columns(sheet1,@user_columns)
        end
      end

      row = sheet1.row(idx+1)
      row.replace []

      #lf_pos = get_value_width(user.id)
      #columns_width[0] = lf_pos unless columns_width[0] >= lf_pos

      @user_columns.each_with_index do |c, j|
        v = if c.is_a?(UserXlsImpexCustomFieldColumn)
          case c.custom_field.field_format
            when "int"
              begin
                Integer(user.custom_value_for(c.custom_field).to_s)
              rescue
                show_value(user.custom_value_for(c.custom_field))
              end
            when "float"
              begin
                Float(user.custom_value_for(c.custom_field).to_s)
              rescue
                show_value(user.custom_value_for(c.custom_field))
              end
            when "date"
              begin
                Date.parse(user.custom_value_for(c.custom_field).to_s)
              rescue
                show_value(user.custom_value_for(c.custom_field))
              end
          else
            show_value(user.custom_value_for(c.custom_field))
          end
        else
          c.value(user)
        end

        value = %w(Time Date Fixnum Float Integer String).include?(v.class.name) ? v : v.to_s

        lf_pos = get_value_width(value)
        columns_width[j] = lf_pos unless columns_width[j] >= lf_pos
        row << value
      end

      idx = idx + 1

    end

    if sheet1
      update_sheet_formatting(sheet1,columns_width)
    else
      sheet1 = book.create_worksheet(:name => l(:label_user_plural))
      sheet1.row(0).replace [l(:label_no_data)]
    end

    xls_stream = StringIO.new('')
    book.write(xls_stream)

    return xls_stream.string
  end

  def init_header_columns(sheet1,columns)

    #columns_width = [1]
    #sheet1.row(0).replace ["#"]

    columns_width = []
    sheet1.row(0).replace []

    columns.each do |c|
      sheet1.row(0) << c.caption
      columns_width << (get_value_width(c.caption)*1.1)
    end
# id
#    sheet1.column(0).default_format = Spreadsheet::Format.new(:number_format => "0")

    opt = Hash.new
    columns.each_with_index do |c, idx|
      width = 0
      opt.clear

      if c.is_a?(UserXlsImpexCustomFieldColumn)
        case c.custom_field.field_format
          when "int"
            opt[:number_format] = "0"
          when "float"
            opt[:number_format] = "0.00"
        end
      end

      sheet1.column(idx).default_format = Spreadsheet::Format.new(opt) unless opt.empty?
      columns_width[idx] = width unless columns_width[idx] >= width
    end

    return columns_width
  end

  def update_sheet_formatting(sheet1,columns_width)

    sheet1.row(0).count.times do |idx|

      do_wrap = columns_width[idx] > 60 ? 1 : 0
      sheet1.column(idx).width = columns_width[idx] > 60 ? 60 : columns_width[idx]

      if do_wrap
        fmt = Marshal::load(Marshal.dump(sheet1.column(idx).default_format))
        fmt.text_wrap = true
        sheet1.column(idx).default_format = fmt
      end

      fmt = Marshal::load(Marshal.dump(sheet1.row(0).format(idx)))
      fmt.font.bold=true
      fmt.pattern=1
      fmt.pattern_bg_color=:gray
      fmt.pattern_fg_color=:gray
      sheet1.row(0).set_format(idx,fmt)
    end

  end

  def get_value_width(value)

    if ['Time', 'Date'].include?(value.class.name)
      return 18 unless value.to_s.length < 18
    end

    tot_w = Array.new
    tot_w << Float(0)
    idx=0
    value.to_s.each_char do |c|
      case c
        when '1', '.', ';', ':', ',', ' ', 'i', 'I', 'j', 'J', '(', ')', '[', ']', '!', '-', 't', 'l'
          tot_w[idx] += 0.6
        when 'W', 'M', 'D'
          tot_w[idx] += 1.2
        when "\n"
          idx = idx + 1
          tot_w << Float(0)
      else
        tot_w[idx] += 1.05
      end
    end

    wdth=0
    tot_w.each do |w|
      wdth = w unless w<wdth
    end

    return wdth+1.5
  end

  def pretty_xls_tab_name(org_name)
    return org_name.gsub(/[\\\/\[\]\?\*:"']/, '_')
  end

end
