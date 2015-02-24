require 'uiex_asset_helpers'

module UsersXlsImpexHelper
  unloadable

  include UsersHelper

  def uxlsie_options_for_list_of_worksheets
    fields = []
    @xls_book.worksheets.each_with_index do |w,i|
      fields << [w.name, i]
    end

    return options_for_select(fields)
  end

# column_idx - integer
  def uxlsie_get_column_options(column_idx,column_name)

    ret_str = "<option value=\"\">#{h(l(:label_users_xlsie_option_ignore))}</option>".html_safe

    selected_name = nil

    if session[:users_xlsie_columns_map] != nil
      xls_columns_map=session[:users_xlsie_columns_map]
      selected_name = xls_columns_map[column_idx.to_s] unless xls_columns_map[column_idx.to_s] == ''
    else
      c=@user_columns.detect do |uc|
        true if uc.caption == column_name
      end
      selected_name = c.name.to_s if c
    end

    ret_str << options_from_collection_for_select(@user_columns, 'name', 'caption', selected_name ? selected_name.to_sym : nil )

    return ret_str.html_safe
  end

# column_idx - integer
  def uxlsie_get_column_update_state(column_idx)
    return false unless session[:users_xlsie_columns_update_map]

    update_opt=session[:users_xlsie_columns_update_map][column_idx.to_s]
    if update_opt
      return true if update_opt == '1'
    end
    return false
  end

  def uxlsie_get_update_options
    options_array = [[l(:label_users_xlsie_opt_update_a),0],[l(:label_users_xlsie_opt_update_b),1],[l(:label_users_xlsie_opt_update_c),2]]

    selected_opt = 0
    selected_opt=session[:users_xlsie_update_only].to_i if session[:users_xlsie_update_only]

    return options_for_select(options_array, selected_opt).html_safe
  end

# column_name - string, column_idx - integer
  def uxlsie_get_column_pars_text(column_name,column_idx)
    txt = ''

    if session[:users_xlsie_columns_pars_map] != nil
      xls_columns_pars_map=session[:users_xlsie_columns_pars_map]
      txt = xls_columns_pars_map[column_idx.to_s] unless xls_columns_pars_map[column_idx.to_s] == ''
    end

    return txt
  end

  def uxlsie_list_of_saved_users
    ret_str = ''

    @users_saved.each_with_index do |is,i|
      ret_str << link_to("#{is[:user].name}", {:controller => "users", :action => "show", :id => is[:user].id })
      ret_str << ', ' unless i == @users_saved.count-1
    end

    return ret_str.html_safe
  end

  def uxlsie_list_of_updated_users
    ret_str = ''

    @users_updated.each_with_index do |is,i|
      ret_str << link_to("#{is.name}", {:controller => "users", :action => "show", :id => is.id })
      ret_str << ', ' unless i == @users_updated.count-1
    end

    return ret_str.html_safe
  end

  def uxlsie_list_of_duplicated_users
    ret_str = ''

    @users_duplicated.each_with_index do |isa,i|
      ret_str << "["
      isa.each_with_index do |is,j|
        ret_str << link_to("#{is.name}", {:controller => "users", :action => "show", :id => is.id})
        ret_str << ', ' unless j == isa.count-1
      end
      ret_str << "]"
      ret_str << ', ' unless i == @users_duplicated.count-1
    end

    return ret_str.html_safe
  end

  def uxlsie_get_preview_row(row_idx)
    ret_str = ''
    row=@xls_book.worksheet(@xls_sheet_num).row(row_idx)

    @xls_columns.each_with_index do |c,idx|
      v=if row[idx]==nil
        ' '
      else
        if row[idx].is_a?(Spreadsheet::Formula)
          '[F] '+row[idx].value.to_s
        else
          row[idx].to_s
        end
      end
      ret_str << content_tag('td', v).html_safe
    end

    return ret_str
  end

  def uxlsie_format_user_errors(ic)
    ret_str = ''
    err_array=ic[:validation_errors]
    if err_array
      err_array.each_with_index do |e,idx|
        ret_str << h(e[:attr_name]) << '(' << e[:message] << ')'
        ret_str << '<br/>' unless idx == err_array.count-1
      end
    end
    return ret_str.html_safe
  end

end
