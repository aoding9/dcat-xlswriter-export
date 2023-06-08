<?php

namespace Aoding9\Dcat\Xlswriter\Export\Demo;

use Dcat\Admin\Traits\HasDateTimeFormatter;
use Illuminate\Contracts\Auth\Authenticatable;
use Illuminate\Database\Eloquent\SoftDeletes;
use Illuminate\Database\Eloquent\Model;

/**
 * App\Models\User
 *
 * @property int $id
 * @property \Illuminate\Support\Carbon|null $created_at
 * @property \Illuminate\Support\Carbon|null $updated_at
 * @property \Illuminate\Support\Carbon|null $deleted_at
 * @property string|null $name 名称
 * @property string|null $username 用户名
 * @property string|null $password 密码
 * @property string|null $phone 手机号
 * @property string|null $number 工号
 * @property int|null $department_id 部门
 * @property int|null $office_area_id 办公区
 * @property int|null $on_blacklist 是否黑名单
 * @property string|null $openid
 * @property string|null $wx_session_key

 */
class User extends Model implements Authenticatable {
    use HasDateTimeFormatter;
    use SoftDeletes;
    use \Illuminate\Auth\Authenticatable;
    
}
