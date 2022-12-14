# ���O / Name

MMVC_Trainer Support VBScript

# �Љ� / Features

MMVC_Trainer�ŋ@�B�w�K����̂ɖ𗧂��Ǝv���A�x���c�[����VBScript�ŏ����Ă݂܂����B

1. katakana.vbs: �e�L�X�g�t�@�C���̒�����J�^�J�i�̎g�p���𒲂ׂă��X�g�A�b�v�B
2. duration.vbs: WAV�t�@�C���̒������͈͓����𒲂ׂă��X�g�A�b�v�B
3. encode.vbs: �����t�@�C�����@�B�w�K�p�ɃG���R�[�h�B


# �K�{���� / Requirement

* VBScript�����s�ł���Windows��
* [ffmpeg](https://ffmpeg.org/) (ffmpeg.exe, ffprobe.exe) 

# �C���X�g�[�����@

�����t�@�C���̑���ɂ�ffmpeg���g�p���܂��̂ŁAffmpeg���_�E�����[�h���āA�X�N���v�g�t�@�C��(.vbs)�Ɠ����t�H���_��ffmpeg.exe��ffprobe.exe��ݒu���Ă��������B

# Usage

**katakana.vbs�̎g����**
    katakana.vbs���_�u���N���b�N���܂��B�utext�v�t�H���_�ɓ����Ă���txt�t�@�C����T�����ăJ�^�J�i�g�p�����W�v�����ukatakana.txt�v����������܂��B
    �g�p����0�������ꍇ�́u���v���t���܂��B

**duration.vbs�̎g����**
    duration.vbs���_�u���N���b�N���܂��B�uwav�v�t�H���_�ɓ����Ă���wav�t�@�C����T�����Ē����𒲂ׂ��uduration.txt�v����������܂��B
    �͈͊O(0.401�b�����܂��́A15.99�b����)�̃t�@�C�����������ꍇ�́u���v���t���܂��B

**encode.vbs�̎g����**
    encode.vbs���_�u���N���b�N���܂��B�usrc�v�t�H���_�ɓ����Ă��鉹���t�@�C��(wav, mp3, ogg)�t�@�C����T�����āuwav�v�t�H���_�ɃG���R�[�h���܂��B
    ���߂�ꂽ�t�H�[�}�b�g(24000Hz 16bit 1ch)��wav�t�@�C���������ꍇ�́A�G���R�[�h���Ȃ���wav�t�@�C���𕡐����܂��B


# Note

�@�B�w�K�̕׋����ł��B�K�v�ȃX�N���v�g������΍�낤�Ǝv���܂��B

# ��� / Author

* xbmc_now
* [@xbmc_now](https://twitter.com/xbmc_now)

# License
