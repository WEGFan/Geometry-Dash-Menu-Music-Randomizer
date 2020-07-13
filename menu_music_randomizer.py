# -*- coding: utf-8 -*-
import locale
import msvcrt
import os
import random
import sys
import time
import traceback
from pathlib import Path
from textwrap import dedent
from typing import Optional, List, Union

import colorama
import pymem
import win32gui
import win32process
from colorama import Fore, Style, Cursor

__program_name__ = 'Geometry Dash Menu Music Randomizer'
__author__ = 'WEGFan'
__version__ = '1.0.0'


def get_multi_level_offset(game: pymem.Pymem, offset_list: List[int]) -> int:
    """Get the address result of [base+A]+B]+...]+C, [X] means the value at address X.

    :param game: a pymem.Pymem object, to load memory and get base address
    :param offset_list: a list contains a sequence of hex offset values
    :return: the address result
    """
    if not isinstance(offset_list, list):
        raise TypeError("offset list must be 'list'")
    if len(offset_list) == 0:
        raise ValueError("offset list must not be empty")
    base_address = game.process_base.lpBaseOfDll
    address = base_address
    for offset in offset_list[:-1]:
        address = game.read_uint(address + offset)
    address += offset_list[-1]
    return address


def get_process_id_by_window(class_: Optional[str], title: Optional[str]) -> Optional[int]:
    """Get process id by windows class and/or title.
    If class or title is None, only search by another parameter.

    :param class_: Window class
    :param title: Window title
    :return: process id if window found, otherwise None
    """
    window = win32gui.FindWindow(class_, title)
    if not window:
        return None
    (_, pid) = win32process.GetWindowThreadProcessId(window)
    return pid


def pymem_hook():
    """Hook Pymem module"""
    # suppress outputs
    pymem.process.print = lambda *args, **kwargs: None
    pymem.logger.setLevel(9999999)

    # fix overflow error on windows 7 caused by process handle, don't know why it works
    @property
    def process_handle(self):
        if isinstance(self._process_handle, int):
            # on windows 7 process_handle returns a 8-byte value, so take low 4 bytes to prevent overflow error
            return self._process_handle & 0xffffffff
        return self._process_handle

    @process_handle.setter
    def process_handle(self, value):
        self._process_handle = value

    pymem.Pymem.process_handle = process_handle


def main():
    pymem_hook()

    os.system(f'title {__program_name__}')
    print(dedent(
        f'''\
        {Fore.CYAN}{__program_name__} v{__version__} by {__author__}

        {Fore.MAGENTA}Randomize menu music every time you return to the menu.

        Usage: Simply copy files to <game installation folder>\\Resources\\Menu Music.
        The program will help you create and open it if it doesn't exist.

        Known issue:
        - Music will reset when returning to the title screen
        '''
    ))
    game = pymem.Pymem()
    while True:
        no_music_message_printed = False  # only print hint message once
        while True:
            pid = get_process_id_by_window('GLFW30', 'Geometry Dash')
            if not pid:
                if not no_music_message_printed:
                    print(f'{Fore.RED}Waiting for Geometry Dash... Please make sure the game is opened.')
                    no_music_message_printed = True
                time.sleep(1)
            else:
                game.open_process_from_id(pid)
                print(f'{Fore.GREEN}Game loaded.')
                break

        # get os encoding to correctly decode ascii characters
        os_encoding = locale.getpreferredencoding()
        game_file_path = Path(game.process_base.filename.decode(os_encoding)).resolve()
        game_file_size = game_file_path.stat().st_size

        # check whether game version is 2.113 by file size
        correct_game_file_size = 6854144
        if game_file_size != correct_game_file_size:
            print(f'{Fore.RED}Your game version is not 2.113 '
                  f'(exe size is {game_file_size} bytes, should be {correct_game_file_size} bytes)')
            print(f'{Fore.RED}Press any key to exit the program.')
            msvcrt.getch()
            sys.exit(1)

        # patch the memory
        offsets = [
            0x24530, 0x24977, 0x249a4, 0xce8a8, 0x14bb1c, 0x1583ec, 0x18cefc,
            0x1907ef, 0x1ddf5c, 0x20d9e2, 0x21f989, 0x22471b, 0x22b308
        ]
        new_address = game.allocate(4 * 1024)
        for offset in offsets:
            address = get_multi_level_offset(game, [offset])
            game.write_uint(address, new_address)
        game.write_string(new_address, 'menuLoop.mp3' + '\x00')

        game_directory = game_file_path.parent
        music_directory = game_directory / 'Resources' / 'Menu Music'
        if not music_directory.exists():
            os.mkdir(music_directory)
            os.startfile(music_directory)

        print(f'{Fore.GREEN}Searching music in {music_directory}')

        no_music_message_printed = False
        music_found_previously = False
        try:
            while True:
                music_files = [item for item in music_directory.glob('*') if item.is_file()]

                if not music_files:
                    if not no_music_message_printed:
                        print(fr'{Fore.RED}There are no music files in Resources\Menu Music directory. '
                              'Menu song restored to default.')
                        no_music_message_printed = True
                    game.write_string(new_address, 'menuLoop.mp3' + '\x00')  # restore to default menu music
                    music_found_previously = False
                else:
                    if music_found_previously:
                        # clear the previous line to make "Found x music" always on the same line
                        print(Cursor.UP() + colorama.ansi.clear_line(), end='')
                    print(f'{Fore.GREEN}Found {len(music_files)} music.')
                    music_file = random.choice(music_files)
                    game.write_string(new_address, str(music_file.resolve()) + '\x00')
                    no_music_message_printed = False
                    music_found_previously = True

                time.sleep(1)
        except pymem.exception.MemoryWriteError as err:
            # check whether exception is caused by game close or unexpected errors
            all_process_id = [process.th32ProcessID for process in pymem.process.list_processes()]
            if game.process_id not in all_process_id:
                continue
            raise
        finally:
            game.close_process()


if __name__ == '__main__':
    try:
        colorama.init(autoreset=False)
        print(Style.BRIGHT, end='')
        main()
    except (KeyboardInterrupt, EOFError) as err:
        sys.exit()
    except Exception as err:
        github_new_issue_url = 'https://github.com/WEGFan/Geometry-Dash-Menu-Music-Randomizer/issues/new'
        print(dedent(
            f'''\
            {Fore.RED}Oops, something went wrong...
            Create an issue on Github ({github_new_issue_url}) with the following yellow lines to let me know what happened!
            '''
        ))
        print(f'{Fore.YELLOW}{traceback.format_exc()}')
        print(f'{Fore.RED}Press any key to exit the program.')
        msvcrt.getch()
        sys.exit(1)
