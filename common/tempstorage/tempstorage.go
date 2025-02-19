//
// Copyright 2020 FoxyUtils ehf. All rights reserved.
//
// This is a commercial product and requires a license to operate.
// A trial license can be obtained at https://unidoc.io
//
// DO NOT EDIT: generated by unitwist Go source code obfuscator.
//
// Use of this source code is governed by the UniDoc End User License Agreement
// terms that can be accessed at https://unidoc.io/eula/

package tempstorage ;import _ae "io";

// TempDir creates a name for a new temp directory using a pattern argument.
func TempDir (pattern string )(string ,error ){return _g .TempDir (pattern )};

// SetAsStorage changes temporary storage to newStorage.
func SetAsStorage (newStorage storage ){_g =newStorage };type storage interface{Open (_f string )(File ,error );TempFile (_d ,_c string )(File ,error );TempDir (_ag string )(string ,error );RemoveAll (_e string )error ;Add (_aa string )error ;};

// Open returns tempstorage File object by name.
func Open (path string )(File ,error ){return _g .Open (path )};

// RemoveAll removes all files according to the dir argument prefix.
func RemoveAll (dir string )error {return _g .RemoveAll (dir )};

// Add reads a file from a disk and adds it to the storage.
func Add (path string )error {return _g .Add (path )};

// File is a representation of a storage file
// with Read, Write, Close and Name methods identical to os.File.
type File interface{_ae .Reader ;_ae .ReaderAt ;_ae .Writer ;_ae .Closer ;Name ()string ;};

// TempFile creates new empty file in the storage and returns it.
func TempFile (dir ,pattern string )(File ,error ){return _g .TempFile (dir ,pattern )};var _g storage ;