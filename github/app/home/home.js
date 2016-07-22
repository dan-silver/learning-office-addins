(function(){
  'use strict';

  // The initialize function must be run each time a new page is loaded
  Office.initialize = function(reason){
    jQuery(document).ready(function(){
      app.initialize();

      jQuery('#search-google').click(function () {
        Office.context.document.getSelectedDataAsync(Office.CoercionType.Text,
          function(result){
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                var win = window.open("https://www.google.com/webhp?sourceid=chrome-instant&ion=1&espv=2&ie=UTF-8#q=" + result.value, '_blank');
                win.focus();
            } else {
              app.showNotification('Error:', result.error.message);
            }
          }
        );
      })
    });
  };
})();
