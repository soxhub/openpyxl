Performance
===========

openpyxl attempts to balance functionality and performance. Where in doubt,
we have focused on functionality over optimisation: performance tweaks are
easier once an API has been established. Memory use is fairly high in
comparison with other libraries and applications and is approximately 50
times the original file size, e.g. 2.5 GB for a 50 MB Excel file. As many use
cases involve either only reading or writing files, the :doc:`optimized`
modes mean this is less of a problem.


Benchmarks
----------

All benchmarks are synthetic and extremely dependent upon the hardware but
they can nevertheless give an indication.


Write Performance
+++++++++++++++++

The `benchmark code
<https://bitbucket.org/snippets/charlie_x/5edaGE/write-performance-benchmark>`_
can be adjusted to use more sheets and adjust the proportion of data that is
strings. Because the version of Python being used can also significantly
affect performance, a `driver script
<https://bitbucket.org/snippets/charlie_x/gebj7M/drive-a-script-through-different-python>`_
can also be used to test with different Python versions with a tox
environment.

Performance is compared with the excellent alternative library xlsxwriter

.. literalinclude:: write_performance.txt


Read Performance
++++++++++++++++

Performance is measured using a file provided with a previous `bug report
<https://bitbucket.org/openpyxl/openpyxl/issues/494/>`_ and compared with the
older xlrd library. xlrd is primarily for the older BIFF file format of .XLS
files but it does have limited support for XLSX.

The code for the `benchmark
<https://bitbucket.org/snippets/openpyxl/Ee9zqo>`_ shows the importance of
choosing the right options when working with a file. In this case disabling
external links.

One major difference between the libraries is that openpyxl's read-only mode
opens a workbook almost immediately making it suitable for multiple
processes.


.. literalinclude:: read_performance.txt
