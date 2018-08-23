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
they can nevertheless give an indication. The `XLSXWriter benchmarks <`_ I've
updated to use more openpyxl more idiomatically and with lxml installed get
these results.

.. literalinclude:: benchmark.txt
