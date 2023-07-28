<fieldset class="form-group">
    <legend class="col-form-label-lg">Lactoscope H230613-16 COMP</legend>
    <label for='instrument_results_file' class="text-muted">
        You can upload an XLS, XLSX, or CSV file.
    </label>
    <input type="file"
           class="form-control-file"
           name="instrument_results_file"
           id="instrument_results_file"/>
</fieldset>

<input name="firstsubmit"
       type="submit"
       value="Submit"
       class="btn btn-primary"
         i18n:attributes="value"/>

<hr/>

<fieldset class="form-group">
    <legend class="col-form-label-lg">Advanced options</legend>

    <div class="form-group">
        <label for="artoapply">Samples state</label>
        <select name="artoapply" class="form-control" id="artoapply">
            <option value="received">
                Received
            </option>
            <option value="received_tobeverified">
                Received and to be verified
            </option>
        </select>
    </div>

    <div class="form-group">
        <label for="results_override">Results override</label>
        <select name="results_override" id="results_override" class="form-control">
            <option value="nooverride">
                Don't override results
            </option>
            <option value="override">
                Override non-empty results
            </option>
            <option value="overrideempty">
                Override non-empty results (also with empty)
            </option>
        </select>
    </div>

</fieldset>
