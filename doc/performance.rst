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
external links stops openpyxl opening cached copies of the linked worksheets.

One major difference between the libraries is that openpyxl's read-only mode
opens a workbook almost immediately making it suitable for multiple
processes, this also readuces memory use significantly. xlrd does also not
automatically convert dates and times into Python datetimes, though it does
annotate cells accordingly but to do this in client code significantly
reduces performance.


.. literalinclude:: read_performance.txt


Parallelisation
+++++++++++++++

Reading worksheets is fairly CPU-intensive which limits any benefits to be
gained by parallelisation. However, if you are mainly interested in dumping
the contents of a workbook then you can use openpyxl's read-only mode and
open multiple instances of a workbook and take advantage of multiple CPUs.

`Sample code <https://bitbucket.org/snippets/openpyxl/AexG8E>`_ using the
same source file as for read performance shows that performance scales
reasonably with only a slight overhead due to creating additional Python
processes.

.. code-block::

    Parallised Read
        Workbook loaded 1.12s
        >>DATA>> 2.27s
        Output Model 2.30s
        Store days 100% 37.18s
        OptimizationData 44.09s
        Store days 0% 45.60s
        Total time 46.76s
