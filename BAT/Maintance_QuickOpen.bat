@echo off
start diskmgmt.msc
start eventvwr.msc /MAX
start control update
del %0